# -*- coding: utf-8 -*-
"""
Recherche Outlook / .eml - Application de recherche avancee multi-criteres
==========================================================================
Petite application avec interface (Tkinter) qui recherche des emails sur
plusieurs criteres en meme temps, depuis deux sources au choix :

  - SOURCE "Outlook" : se connecte a Outlook bureau. On peut laisser le
    compte par defaut, ou taper l'adresse d'une boite precise (la sienne
    ou une boite partagee) pour explorer SES dossiers.

  - SOURCE "Fichiers .eml" : recherche dans un dossier local contenant
    des fichiers .eml (avec ou sans les sous-dossiers). Ne necessite PAS
    Outlook.

Criteres : mots-cles (ET / OU) dans l'objet et/ou le corps, expediteur,
objet contient, corps contient, plage de dates, presence de piece jointe.

La configuration (criteres + source) peut etre ENREGISTREE puis rechargee
automatiquement au demarrage.

PREREQUIS :
  - Mode Outlook : Windows + Outlook bureau + pywin32 (pip install pywin32)
  - Mode .eml    : Python seul (aucune dependance)
LANCEMENT : python recherche_outlook.py
"""

import os
import sys
import json
import queue
import csv
import threading
import datetime as dt
import email
import email.utils
from email import policy
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import win32com.client
    import pythoncom
    PYWIN32_OK = True
except ImportError:
    PYWIN32_OK = False


OL_MAIL_CLASS = 43      # olMail
OL_FOLDER_INBOX = 6     # olFolderInbox
CONFIG_NAME = "config_recherche_outlook.json"


# --------------------------------------------------------------------------
#  Emplacement du fichier de configuration
# --------------------------------------------------------------------------
def config_path():
    try:
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        base = os.path.expanduser("~")
    return os.path.join(base, CONFIG_NAME)


# --------------------------------------------------------------------------
#  Filtrage commun (s'applique a un "row" deja extrait)
# --------------------------------------------------------------------------
def text_matches(subject, body, sender_all, criteria):
    subject_l = (subject or "").lower()
    body_l = (body or "").lower()
    sender_all = (sender_all or "").lower()

    if criteria["keywords"]:
        if criteria["scope"] == "objet":
            target = subject_l
        elif criteria["scope"] == "corps":
            target = body_l
        else:
            target = subject_l + " " + body_l
        kws = criteria["keywords"]
        if criteria["mode"] == "ET":
            if not all(k in target for k in kws):
                return False
        else:
            if not any(k in target for k in kws):
                return False

    if criteria["sender"] and criteria["sender"] not in sender_all:
        return False
    if criteria["subject"] and criteria["subject"] not in subject_l:
        return False
    if criteria["body"] and criteria["body"] not in body_l:
        return False
    return True


def date_in_range(when, criteria):
    """when : datetime ou None."""
    if when is None:
        return not (criteria["date_from"] or criteria["date_to"])
    if criteria["date_from"] and when.date() < criteria["date_from"].date():
        return False
    if criteria["date_to"] and when.date() > criteria["date_to"].date():
        return False
    return True


# --------------------------------------------------------------------------
#  Source Outlook
# --------------------------------------------------------------------------
def build_date_filter(date_from, date_to):
    parts = []
    fmt = "%m/%d/%Y %H:%M %p"
    if date_from:
        parts.append("\"urn:schemas:httpmail:datereceived\" >= '%s'"
                     % date_from.strftime(fmt))
    if date_to:
        end = date_to + dt.timedelta(days=1)
        parts.append("\"urn:schemas:httpmail:datereceived\" < '%s'"
                     % end.strftime(fmt))
    if not parts:
        return None
    return "@SQL=" + " AND ".join(parts)


def ol_item_to_row(item, criteria):
    try:
        if item.Class != OL_MAIL_CLASS:
            return None
    except Exception:
        return None
    try:
        subject = item.Subject or ""
    except Exception:
        subject = ""
    try:
        body = item.Body or ""
    except Exception:
        body = ""
    try:
        sender = item.SenderName or ""
    except Exception:
        sender = ""
    try:
        sender_email = item.SenderEmailAddress or ""
    except Exception:
        sender_email = ""

    if not text_matches(subject, body, sender + " " + sender_email, criteria):
        return None

    try:
        nb_pj = item.Attachments.Count
    except Exception:
        nb_pj = 0
    if criteria["has_attach"] and nb_pj == 0:
        return None

    try:
        received_str = item.ReceivedTime.strftime("%Y-%m-%d %H:%M")
    except Exception:
        received_str = ""
    try:
        entry_id = item.EntryID
    except Exception:
        entry_id = ""
    try:
        store_id = item.Parent.StoreID
    except Exception:
        store_id = ""

    return {"date": received_str, "from": sender or sender_email,
            "subject": subject, "pj": "Oui (%d)" % nb_pj if nb_pj else "",
            "source": "outlook", "entry_id": entry_id, "store_id": store_id,
            "path": ""}


class OutlookWorker(threading.Thread):
    def __init__(self, in_queue, out_queue, stop_event, mailbox_email=""):
        super().__init__(daemon=True)
        self.inq = in_queue
        self.outq = out_queue
        self.stop_event = stop_event
        self.mailbox_email = (mailbox_email or "").strip()
        self.app = None
        self.ns = None
        self.folders = []

    def _connect(self):
        self.outq.put(("status", "Demarrage d'Outlook..."))
        self.app = None
        try:
            self.app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            self.app = None
        if self.app is None:
            self.app = win32com.client.Dispatch("Outlook.Application")

        self.outq.put(("status", "Lecture des dossiers Outlook..."))
        self.ns = self.app.GetNamespace("MAPI")
        self.folders = self._list_folders()
        if not self.folders:
            self.outq.put(("error",
                           "Connexion etablie mais aucun dossier trouve. "
                           "Verifiez l'adresse saisie ou le compte Outlook."))
            return
        self.outq.put(("folders", [p for p, _ in self.folders]))

    def _walk(self, folder, prefix, out):
        try:
            name = folder.Name
        except Exception:
            return
        path = prefix + name
        out.append((path, folder))
        try:
            subs = folder.Folders
        except Exception:
            return
        try:
            for sub in subs:
                self._walk(sub, path + " / ", out)
        except Exception:
            pass

    def _list_folders(self):
        out = []
        try:
            if self.mailbox_email:
                # Boite precise (la sienne ou partagee)
                recip = self.ns.CreateRecipient(self.mailbox_email)
                recip.Resolve()
                if not recip.Resolved:
                    self.outq.put(("error",
                                   "Adresse introuvable : " + self.mailbox_email))
                    return []
                inbox = self.ns.GetSharedDefaultFolder(recip, OL_FOLDER_INBOX)
                try:
                    root = inbox.Store.GetRootFolder()
                    self._walk(root, "", out)
                except Exception:
                    # a defaut, au moins la boite de reception
                    self._walk(inbox, "", out)
            else:
                for store in self.ns.Folders:
                    self._walk(store, "", out)
        except Exception as exc:
            self.outq.put(("error", "Lecture des dossiers : " + str(exc)))
        return out

    def _search(self, folder_index, criteria, max_results):
        try:
            folder = self.folders[folder_index][1]
            items = folder.Items
            try:
                items.Sort("[ReceivedTime]", True)
            except Exception:
                pass
            df = build_date_filter(criteria["date_from"], criteria["date_to"])
            if df:
                try:
                    items = items.Restrict(df)
                except Exception:
                    pass
            scanned = found = 0
            item = items.GetFirst()
            while item is not None:
                if self.stop_event.is_set():
                    break
                scanned += 1
                if scanned % 50 == 0:
                    self.outq.put(("progress", scanned, found))
                row = ol_item_to_row(item, criteria)
                if row is not None:
                    self.outq.put(("row", row))
                    found += 1
                    if found >= max_results:
                        self.outq.put(("limit", max_results))
                        break
                item = items.GetNext()
            self.outq.put(("done", scanned, found))
        except Exception as exc:
            self.outq.put(("error", "Recherche : " + str(exc)))

    def _open(self, entry_id, store_id):
        try:
            item = self.ns.GetItemFromID(entry_id, store_id)
            item.Display()
        except Exception as exc:
            self.outq.put(("error", "Ouverture du mail : " + str(exc)))

    def run(self):
        pythoncom.CoInitialize()
        try:
            self._connect()
        except Exception as exc:
            self.outq.put(("error", "Connexion a Outlook : " + str(exc)))
            pythoncom.CoUninitialize()
            return
        try:
            while True:
                cmd = self.inq.get()
                if cmd[0] == "quit":
                    break
                elif cmd[0] == "search":
                    self.stop_event.clear()
                    self._search(cmd[1], cmd[2], cmd[3])
                elif cmd[0] == "open":
                    self._open(cmd[1], cmd[2])
        finally:
            pythoncom.CoUninitialize()


# --------------------------------------------------------------------------
#  Source fichiers .eml (local, sans Outlook)
# --------------------------------------------------------------------------
def eml_get_body(msg):
    try:
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain" and \
                        part.get_content_disposition() != "attachment":
                    try:
                        return part.get_content()
                    except Exception:
                        payload = part.get_payload(decode=True)
                        if payload:
                            return payload.decode(errors="replace")
            return ""
        else:
            try:
                return msg.get_content()
            except Exception:
                payload = msg.get_payload(decode=True)
                return payload.decode(errors="replace") if payload else ""
    except Exception:
        return ""


def eml_count_attachments(msg):
    n = 0
    try:
        for part in msg.walk():
            if part.get_content_disposition() == "attachment" or \
                    part.get_filename():
                n += 1
    except Exception:
        pass
    return n


def eml_to_row(path, criteria):
    try:
        with open(path, "rb") as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
    except Exception:
        return None

    subject = str(msg.get("subject", "") or "")
    sender = str(msg.get("from", "") or "")
    body = eml_get_body(msg)

    if not text_matches(subject, body, sender, criteria):
        return None

    when = None
    raw_date = msg.get("date")
    if raw_date:
        try:
            when = email.utils.parsedate_to_datetime(raw_date)
            if when is not None and when.tzinfo is not None:
                when = when.replace(tzinfo=None)
        except Exception:
            when = None
    if not date_in_range(when, criteria):
        return None

    nb_pj = eml_count_attachments(msg)
    if criteria["has_attach"] and nb_pj == 0:
        return None

    date_str = when.strftime("%Y-%m-%d %H:%M") if when else ""
    return {"date": date_str, "from": sender, "subject": subject,
            "pj": "Oui (%d)" % nb_pj if nb_pj else "",
            "source": "eml", "entry_id": "", "store_id": "", "path": path}


def list_eml_files(folder, recursive):
    files = []
    if recursive:
        for root, _dirs, names in os.walk(folder):
            for n in names:
                if n.lower().endswith(".eml"):
                    files.append(os.path.join(root, n))
    else:
        try:
            for n in os.listdir(folder):
                if n.lower().endswith(".eml"):
                    files.append(os.path.join(folder, n))
        except Exception:
            pass
    return files


def search_eml(folder, recursive, criteria, outq, stop_event, max_results):
    try:
        files = list_eml_files(folder, recursive)
        if not files:
            outq.put(("error", "Aucun fichier .eml trouve dans ce dossier."))
            outq.put(("done", 0, 0))
            return
        scanned = found = 0
        for path in files:
            if stop_event.is_set():
                break
            scanned += 1
            if scanned % 25 == 0:
                outq.put(("progress", scanned, found))
            row = eml_to_row(path, criteria)
            if row is not None:
                outq.put(("row", row))
                found += 1
                if found >= max_results:
                    outq.put(("limit", max_results))
                    break
        outq.put(("done", scanned, found))
    except Exception as exc:
        outq.put(("error", "Recherche .eml : " + str(exc)))


# --------------------------------------------------------------------------
#  Interface graphique
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Recherche emails - Outlook / .eml")
        self.geometry("1040x720")
        self.minsize(860, 600)

        self.outq = queue.Queue()
        self.inq = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.results = []
        self._sort_state = {}
        self.folder_paths = []
        self._connected = False
        self._saved_folder = ""

        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(150, self._poll_queue)

        self.load_config(silent=True)
        self._update_source_state()

    # ---------------- UI ----------------
    def _build_ui(self):
        pad = {"padx": 6, "pady": 4}

        # ----- Source -----
        src = ttk.LabelFrame(self, text="Source des emails")
        src.pack(fill="x", padx=10, pady=(8, 4))

        self.var_source = tk.StringVar(value="outlook")
        ttk.Radiobutton(src, text="Outlook", value="outlook",
                        variable=self.var_source,
                        command=self._update_source_state).grid(
            row=0, column=0, sticky="w", **pad)
        ttk.Radiobutton(src, text="Fichiers .eml (dossier local)",
                        value="eml", variable=self.var_source,
                        command=self._update_source_state).grid(
            row=0, column=1, sticky="w", **pad)

        # Ligne Outlook
        ttk.Label(src, text="Boite (email, optionnel) :").grid(
            row=1, column=0, sticky="e", **pad)
        self.var_mailbox = tk.StringVar()
        self.ent_mailbox = ttk.Entry(src, textvariable=self.var_mailbox, width=34)
        self.ent_mailbox.grid(row=1, column=1, sticky="we", **pad)
        self.btn_reconnect = ttk.Button(src, text="Connecter / Reconnecter",
                                        command=self.start_worker)
        self.btn_reconnect.grid(row=1, column=2, sticky="w", **pad)

        ttk.Label(src, text="Dossier Outlook :").grid(
            row=2, column=0, sticky="e", **pad)
        self.var_folder = tk.StringVar()
        self.cb_folder = ttk.Combobox(src, textvariable=self.var_folder,
                                      width=60, state="readonly")
        self.cb_folder.grid(row=2, column=1, columnspan=2, sticky="we", **pad)

        # Ligne .eml
        ttk.Label(src, text="Dossier .eml :").grid(
            row=3, column=0, sticky="e", **pad)
        self.var_eml = tk.StringVar()
        self.ent_eml = ttk.Entry(src, textvariable=self.var_eml, width=60)
        self.ent_eml.grid(row=3, column=1, sticky="we", **pad)
        self.btn_browse = ttk.Button(src, text="Parcourir...",
                                     command=self.browse_eml)
        self.btn_browse.grid(row=3, column=2, sticky="w", **pad)
        self.var_recursive = tk.BooleanVar(value=True)
        self.chk_recursive = ttk.Checkbutton(
            src, text="Inclure les sous-dossiers", variable=self.var_recursive)
        self.chk_recursive.grid(row=4, column=1, sticky="w", **pad)

        for c in range(3):
            src.columnconfigure(c, weight=1)

        # ----- Criteres -----
        crit = ttk.LabelFrame(self, text="Criteres de recherche")
        crit.pack(fill="x", padx=10, pady=4)

        ttk.Label(crit, text="Mots-cles :").grid(row=0, column=0, sticky="e", **pad)
        self.var_keywords = tk.StringVar()
        ttk.Entry(crit, textvariable=self.var_keywords, width=40).grid(
            row=0, column=1, columnspan=2, sticky="we", **pad)
        ttk.Label(crit, text="Combinaison :").grid(row=0, column=3, sticky="e", **pad)
        self.var_mode = tk.StringVar(value="ET")
        ttk.Combobox(crit, textvariable=self.var_mode, values=["ET", "OU"],
                     width=5, state="readonly").grid(row=0, column=4, sticky="w", **pad)
        ttk.Label(crit, text="Chercher dans :").grid(row=0, column=5, sticky="e", **pad)
        self.var_scope = tk.StringVar(value="objet + corps")
        ttk.Combobox(crit, textvariable=self.var_scope,
                     values=["objet + corps", "objet", "corps"],
                     width=12, state="readonly").grid(row=0, column=6, sticky="w", **pad)

        ttk.Label(crit, text="Expediteur :").grid(row=1, column=0, sticky="e", **pad)
        self.var_sender = tk.StringVar()
        ttk.Entry(crit, textvariable=self.var_sender, width=30).grid(
            row=1, column=1, columnspan=2, sticky="we", **pad)
        ttk.Label(crit, text="Objet contient :").grid(row=1, column=3, sticky="e", **pad)
        self.var_subject = tk.StringVar()
        ttk.Entry(crit, textvariable=self.var_subject, width=24).grid(
            row=1, column=4, columnspan=3, sticky="we", **pad)

        ttk.Label(crit, text="Corps contient :").grid(row=2, column=0, sticky="e", **pad)
        self.var_body = tk.StringVar()
        ttk.Entry(crit, textvariable=self.var_body, width=30).grid(
            row=2, column=1, columnspan=2, sticky="we", **pad)
        self.var_attach = tk.BooleanVar(value=False)
        ttk.Checkbutton(crit, text="Uniquement avec piece jointe",
                        variable=self.var_attach).grid(
            row=2, column=3, columnspan=2, sticky="w", **pad)

        ttk.Label(crit, text="Du (AAAA-MM-JJ) :").grid(row=3, column=0, sticky="e", **pad)
        self.var_date_from = tk.StringVar()
        ttk.Entry(crit, textvariable=self.var_date_from, width=14).grid(
            row=3, column=1, sticky="w", **pad)
        ttk.Label(crit, text="Au (AAAA-MM-JJ) :").grid(row=3, column=3, sticky="e", **pad)
        self.var_date_to = tk.StringVar()
        ttk.Entry(crit, textvariable=self.var_date_to, width=14).grid(
            row=3, column=4, sticky="w", **pad)
        ttk.Label(crit, text="Max resultats :").grid(row=3, column=5, sticky="e", **pad)
        self.var_max = tk.StringVar(value="500")
        ttk.Entry(crit, textvariable=self.var_max, width=7).grid(
            row=3, column=6, sticky="w", **pad)

        for c in range(7):
            crit.columnconfigure(c, weight=1)

        # ----- Actions -----
        actions = ttk.Frame(self)
        actions.pack(fill="x", padx=10)
        self.btn_search = ttk.Button(actions, text="Rechercher",
                                     command=self.start_search)
        self.btn_search.pack(side="left", padx=4, pady=4)
        self.btn_stop = ttk.Button(actions, text="Arreter",
                                   command=self.stop_search, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        self.btn_export = ttk.Button(actions, text="Exporter (CSV)",
                                     command=self.export_csv, state="disabled")
        self.btn_export.pack(side="left", padx=4)
        ttk.Button(actions, text="Enregistrer config",
                   command=self.save_config).pack(side="left", padx=4)
        ttk.Button(actions, text="Charger config",
                   command=lambda: self.load_config(silent=False)).pack(
            side="left", padx=4)
        ttk.Button(actions, text="Effacer criteres",
                   command=self.clear_criteria).pack(side="left", padx=4)

        # ----- Resultats -----
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, padx=10, pady=8)
        cols = ("date", "from", "subject", "pj")
        headers = {"date": "Date", "from": "Expediteur",
                   "subject": "Objet", "pj": "PJ"}
        widths = {"date": 130, "from": 220, "subject": 500, "pj": 70}
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=headers[c],
                              command=lambda col=c: self.sort_by(col))
            self.tree.column(c, width=widths[c], anchor="w")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self.open_mail)

        self.var_status = tk.StringVar(value="Pret.")
        ttk.Label(self, textvariable=self.var_status, relief="sunken",
                  anchor="w").pack(fill="x", side="bottom")

    # ---------------- Activer/desactiver selon la source ----------------
    def _update_source_state(self):
        is_ol = self.var_source.get() == "outlook"
        ol_state = "normal" if is_ol else "disabled"
        eml_state = "normal" if not is_ol else "disabled"
        self.ent_mailbox.configure(state=ol_state)
        self.btn_reconnect.configure(state=ol_state)
        self.cb_folder.configure(state="readonly" if is_ol else "disabled")
        self.ent_eml.configure(state=eml_state)
        self.btn_browse.configure(state=eml_state)
        self.chk_recursive.configure(state=eml_state)

    # ---------------- Connexion Outlook ----------------
    def start_worker(self):
        if not PYWIN32_OK:
            messagebox.showerror(
                "Module manquant",
                "Le module 'pywin32' n'est pas installe.\n\n"
                "Tapez dans une invite de commande :\n    pip install pywin32")
            return
        self._connected = False
        self.btn_search["state"] = "normal"
        self.var_status.set("Connexion a Outlook en cours...")
        self.inq = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = OutlookWorker(self.inq, self.outq, self.stop_event,
                                    self.var_mailbox.get())
        self.worker.start()
        self._watchdog_id = self.after(25000, self._connection_watchdog)

    def _connection_watchdog(self):
        if not self._connected and self.var_source.get() == "outlook":
            self.var_status.set("Connexion trop longue - voir l'aide.")
            messagebox.showwarning(
                "Connexion a Outlook",
                "La connexion a Outlook n'aboutit pas.\n\n"
                "1) Une fenetre Outlook est peut-etre cachee (choix de "
                "profil ou message de securite) : cherchez-la dans la barre "
                "des taches et validez-la.\n\n"
                "2) Le \"nouveau Outlook\" de Windows ne permet pas cette "
                "recherche : il faut le Outlook classique (Office / 365).\n\n"
                "Ouvrez Outlook, attendez qu'il soit pret, puis cliquez sur "
                "\"Connecter / Reconnecter\". Vous pouvez aussi utiliser le "
                "mode \"Fichiers .eml\".")

    # ---------------- File des messages ----------------
    def _poll_queue(self):
        try:
            while True:
                msg = self.outq.get_nowait()
                kind = msg[0]
                if kind == "status":
                    self.var_status.set(msg[1])
                elif kind == "folders":
                    self._on_folders(msg[1])
                elif kind == "row":
                    self._add_row(msg[1])
                elif kind == "progress":
                    self.var_status.set(
                        "Analyse... %d examines, %d trouves." % (msg[1], msg[2]))
                elif kind == "limit":
                    self.var_status.set("Limite de %d resultats atteinte." % msg[1])
                elif kind == "done":
                    self._search_finished(msg[1], msg[2])
                elif kind == "error":
                    self.var_status.set("Erreur : " + msg[1])
                    messagebox.showerror("Erreur", msg[1])
                    self._reset_buttons()
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _on_folders(self, paths):
        self._connected = True
        self.folder_paths = paths
        self.cb_folder["values"] = paths
        idx = 0
        if self._saved_folder and self._saved_folder in paths:
            idx = paths.index(self._saved_folder)
        else:
            for i, p in enumerate(paths):
                low = p.lower()
                if "reception" in low or "réception" in low or "inbox" in low:
                    idx = i
                    break
        if paths:
            self.cb_folder.current(idx)
        self.var_status.set("Connecte - %d dossiers disponibles." % len(paths))

    # ---------------- Recherche ----------------
    def _collect_criteria(self):
        kws = [k.strip().lower() for k in self.var_keywords.get().split()
               if k.strip()]
        scope_map = {"objet + corps": "tout", "objet": "objet", "corps": "corps"}
        date_from = self._parse_date(self.var_date_from.get())
        date_to = self._parse_date(self.var_date_to.get())
        if self.var_date_from.get().strip() and date_from is None:
            messagebox.showwarning("Date", "Date 'Du' invalide (AAAA-MM-JJ).")
            return None
        if self.var_date_to.get().strip() and date_to is None:
            messagebox.showwarning("Date", "Date 'Au' invalide (AAAA-MM-JJ).")
            return None
        return {
            "keywords": kws, "mode": self.var_mode.get(),
            "scope": scope_map.get(self.var_scope.get(), "tout"),
            "sender": self.var_sender.get().strip().lower(),
            "subject": self.var_subject.get().strip().lower(),
            "body": self.var_body.get().strip().lower(),
            "has_attach": self.var_attach.get(),
            "date_from": date_from, "date_to": date_to,
        }

    def start_search(self):
        criteria = self._collect_criteria()
        if criteria is None:
            return
        try:
            max_results = int(self.var_max.get())
        except ValueError:
            max_results = 500

        self.tree.delete(*self.tree.get_children())
        self.results = []
        self.stop_event = threading.Event()
        self.btn_search["state"] = "disabled"
        self.btn_stop["state"] = "normal"
        self.btn_export["state"] = "disabled"

        if self.var_source.get() == "outlook":
            if not self._connected or self.worker is None:
                messagebox.showinfo(
                    "Outlook",
                    "Cliquez d'abord sur \"Connecter / Reconnecter\".")
                self._reset_buttons()
                return
            idx = self.cb_folder.current()
            if idx < 0:
                messagebox.showwarning("Dossier", "Choisissez un dossier.")
                self._reset_buttons()
                return
            self.var_status.set("Recherche Outlook en cours...")
            self.inq.put(("search", idx, criteria, max_results))
        else:
            folder = self.var_eml.get().strip()
            if not folder or not os.path.isdir(folder):
                messagebox.showwarning("Dossier .eml",
                                       "Choisissez un dossier valide.")
                self._reset_buttons()
                return
            self.var_status.set("Recherche dans les fichiers .eml...")
            t = threading.Thread(
                target=search_eml,
                args=(folder, self.var_recursive.get(), criteria,
                      self.outq, self.stop_event, max_results),
                daemon=True)
            t.start()

    def stop_search(self):
        self.stop_event.set()
        self.var_status.set("Arret demande...")

    def _add_row(self, row):
        self.results.append(row)
        self.tree.insert("", "end", values=(
            row["date"], row["from"], row["subject"], row["pj"]))

    def _search_finished(self, scanned, found):
        self.var_status.set(
            "Termine : %d examines, %d resultats." % (scanned, found))
        self._reset_buttons()
        if found:
            self.btn_export["state"] = "normal"

    def _reset_buttons(self):
        self.btn_search["state"] = "normal"
        self.btn_stop["state"] = "disabled"

    # ---------------- Tri ----------------
    def sort_by(self, col):
        reverse = self._sort_state.get(col, False)
        self.results.sort(key=lambda r: (r[col] or "").lower(), reverse=reverse)
        self._sort_state[col] = not reverse
        self.tree.delete(*self.tree.get_children())
        for row in self.results:
            self.tree.insert("", "end", values=(
                row["date"], row["from"], row["subject"], row["pj"]))

    # ---------------- Ouverture ----------------
    def open_mail(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx >= len(self.results):
            return
        row = self.results[idx]
        if row["source"] == "outlook":
            if self.worker is not None:
                self.inq.put(("open", row["entry_id"], row["store_id"]))
        else:
            try:
                os.startfile(row["path"])  # Windows : ouvre le .eml
            except Exception as exc:
                messagebox.showerror("Ouverture", str(exc))

    # ---------------- Export ----------------
    def export_csv(self):
        if not self.results:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv")],
            initialfile="resultats_emails.csv")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["Date", "Expediteur", "Objet",
                                 "Piece jointe", "Fichier"])
                for r in self.results:
                    writer.writerow([r["date"], r["from"], r["subject"],
                                     r["pj"], r.get("path", "")])
            self.var_status.set("Export termine : " + path)
            messagebox.showinfo("Export", "Fichier enregistre :\n" + path)
        except Exception as exc:
            messagebox.showerror("Export", str(exc))

    # ---------------- Configuration ----------------
    def save_config(self):
        data = {
            "source": self.var_source.get(),
            "mailbox": self.var_mailbox.get(),
            "outlook_folder": self.var_folder.get(),
            "eml_folder": self.var_eml.get(),
            "eml_recursive": self.var_recursive.get(),
            "keywords": self.var_keywords.get(),
            "mode": self.var_mode.get(),
            "scope": self.var_scope.get(),
            "sender": self.var_sender.get(),
            "subject": self.var_subject.get(),
            "body": self.var_body.get(),
            "has_attach": self.var_attach.get(),
            "date_from": self.var_date_from.get(),
            "date_to": self.var_date_to.get(),
            "max": self.var_max.get(),
        }
        try:
            with open(config_path(), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.var_status.set("Configuration enregistree.")
            messagebox.showinfo("Configuration",
                                "Configuration enregistree :\n" + config_path())
        except Exception as exc:
            messagebox.showerror("Configuration", str(exc))

    def load_config(self, silent=True):
        p = config_path()
        if not os.path.exists(p):
            if not silent:
                messagebox.showinfo("Configuration",
                                    "Aucune configuration enregistree.")
            return
        try:
            with open(p, "r", encoding="utf-8") as f:
                d = json.load(f)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Configuration", str(exc))
            return
        self.var_source.set(d.get("source", "outlook"))
        self.var_mailbox.set(d.get("mailbox", ""))
        self._saved_folder = d.get("outlook_folder", "")
        self.var_eml.set(d.get("eml_folder", ""))
        self.var_recursive.set(d.get("eml_recursive", True))
        self.var_keywords.set(d.get("keywords", ""))
        self.var_mode.set(d.get("mode", "ET"))
        self.var_scope.set(d.get("scope", "objet + corps"))
        self.var_sender.set(d.get("sender", ""))
        self.var_subject.set(d.get("subject", ""))
        self.var_body.set(d.get("body", ""))
        self.var_attach.set(d.get("has_attach", False))
        self.var_date_from.set(d.get("date_from", ""))
        self.var_date_to.set(d.get("date_to", ""))
        self.var_max.set(d.get("max", "500"))
        self._update_source_state()
        # Pre-remplir la liste avec le dossier memorise
        if self._saved_folder:
            self.cb_folder["values"] = [self._saved_folder]
            self.cb_folder.set(self._saved_folder)
        if not silent:
            self.var_status.set("Configuration chargee.")

    # ---------------- Divers ----------------
    def browse_eml(self):
        folder = filedialog.askdirectory(title="Choisir le dossier de .eml")
        if folder:
            self.var_eml.set(folder)

    def clear_criteria(self):
        for var in (self.var_keywords, self.var_sender, self.var_subject,
                    self.var_body, self.var_date_from, self.var_date_to):
            var.set("")
        self.var_attach.set(False)
        self.var_mode.set("ET")
        self.var_scope.set("objet + corps")

    def on_close(self):
        try:
            self.stop_event.set()
            if self.worker is not None:
                self.inq.put(("quit",))
        except Exception:
            pass
        self.destroy()

    @staticmethod
    def _parse_date(text):
        text = text.strip()
        if not text:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        return None


if __name__ == "__main__":
    app = App()
    app.mainloop()
