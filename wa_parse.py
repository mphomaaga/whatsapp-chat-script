#!/usr/bin/env python3
"""
wa_parse.py - turn WhatsApp chats into Excel (.xlsx) and CSV files, offline.

Runs anywhere Python 3.6+ runs: Windows, Mac, Linux, Android (Pydroid 3) and
iPhone (a-Shell). Uses only Python's standard library, so nothing to install.

Where chats can come from:
  - Chat exports from the Android app (.txt) or the iPhone app (.zip)
  - Zips holding several exports (Google Drive / Gmail "download all")
  - WhatsApp Desktop on a Mac: reads its local database, every chat at once
WhatsApp Web and WhatsApp Desktop on Windows have no export option and no
readable database, so for those, export the chats from your phone.

How to run it:
  python wa_parse.py                     no arguments, or just double-click it:
                                         reads the "exports" folder next to this
                                         script, saves to "parsed" next to it
  python wa_parse.py chat.txt            one export
  python wa_parse.py exports             a folder of exports
  python wa_parse.py "exports/*.txt"     wildcards work on Windows too
  python3 wa_parse.py --mac-app          WhatsApp Desktop on this Mac (quit it first)
  python wa_parse.py --mac-app ChatStorage.sqlite   a copy of that database, any OS

What you get: one folder per run, named with the date and time it ran, and
inside it one folder per chat, named after the person (or their number if
they're not saved in your contacts):
  parsed/2026-11-10 09h42/
      Thabo Mokoena/Thabo Mokoena.xlsx    open this one in Excel
      Thabo Mokoena/Thabo Mokoena.csv     same data, for other tools
      +27 82 123 4567/...                 a chat with an unsaved number
"""
import argparse
import csv
import glob
import hashlib
import io
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import traceback
import unicodedata
import zipfile
from collections import Counter
from datetime import datetime
from pathlib import Path, PurePosixPath
from xml.sax.saxutils import escape

# Invisible characters WhatsApp sprinkles into exports (written as chr() so this file stays plain text)
INVISIBLE = {chr(0xFEFF): "", chr(0x200E): "", chr(0x200F): "", chr(0x202A): "", chr(0x202C): "",
             chr(0x202F): " ", chr(0x00A0): " "}

DATE = r"(?P<date>\d{1,4}[./-]\d{1,2}[./-]\d{1,4})"
TIME = r"(?P<time>\d{1,2}:\d{2}(?::\d{2})?(?:\s?[ap]\.?\s?m\.?)?)"
ANDROID = re.compile(r"^" + DATE + r",?\s+" + TIME + r"\s+-\s+(?P<rest>.*)$", re.I)  # 10/11/2026, 08:32 - Ryan: hi
IPHONE = re.compile(r"^\[" + DATE + r",?\s+" + TIME + r"\]\s+(?P<rest>.*)$", re.I)   # [2026/11/10, 08:32:15] Ryan: hi

NAME_PREFIXES = ("whatsapp chat with ", "whatsapp chat - ", "whatsapp chat " + chr(0x2013) + " ")
ORDER_LABELS = {"dmy": "day/month/year", "mdy": "month/day/year", "ymd": "year/month/day"}
COLUMNS = ["Date", "Time", "Sender", "Message"]
WIDTHS = [12, 8, 26, 90]
EXTS = (".txt", ".zip")
EXCEL_MAX_ROWS = 1048575
MAX_NAME = 50                  # keeps folder + file paths well under Windows' 260-character limit
WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL"} | {p + str(i) for p in ("COM", "LPT") for i in range(1, 10)}
XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f" + chr(0xFFFE) + chr(0xFFFF) + "]")

# WhatsApp Desktop for Mac keeps every chat in this SQLite file (same layout as on iPhone)
MAC_DB = Path.home() / "Library" / "Group Containers" / "group.net.whatsapp.WhatsApp.shared" / "ChatStorage.sqlite"
APPLE_EPOCH = 978307200        # its timestamps count seconds from 1 Jan 2001
MEDIA_TYPES = {1: "image", 2: "video", 3: "audio", 4: "contact card", 5: "location",
               8: "document", 11: "GIF", 15: "sticker"}
SKIP_CHATS = ("status@broadcast",)   # status updates aren't conversations
SKIP_SUFFIXES = ("@newsletter",)     # nor are channels you follow


# ---------- reading one exported chat ----------

def clean(s):
    for k, v in INVISIBLE.items():
        s = s.replace(k, v)
    return s


def decode(data):
    """Exports are UTF-8; also cope with one re-saved as UTF-16 ("Unicode") by Notepad."""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return clean(data.decode("utf-16", errors="replace"))
    return clean(data.decode("utf-8-sig", errors="replace"))


def match(line):
    return ANDROID.match(line) or IPHONE.match(line)


def guess_order(dates, forced=None):
    """Work out 'ymd', 'dmy' or 'mdy' from all the date strings in one chat."""
    parts = [re.split(r"[./-]", d) for d in dates]
    if all(len(p[0]) == 4 for p in parts):
        return "ymd"                   # year-first is unambiguous, ignore any flag
    if forced:
        return forced
    if any(int(p[0]) > 12 for p in parts):
        return "dmy"
    if any(int(p[1]) > 12 for p in parts):
        return "mdy"
    return "dmy"  # can't tell: assume day first, like South Africa and most of the world


def to_iso(date, time, order):
    """ISO timestamp, or None if the date doesn't make sense in that order."""
    try:
        a, b, c = (int(x) for x in re.split(r"[./-]", date))
        y, m, d = {"ymd": (a, b, c), "dmy": (c, b, a), "mdy": (c, a, b)}[order]
        if y < 100:
            y += 2000
        t = time.lower().replace(".", "").replace(" ", "")
        ampm = t[-2:] if t.endswith(("am", "pm")) else None
        hms = [int(x) for x in (t[:-2] if ampm else t).split(":")]
        h, mi = hms[0], hms[1]
        s = hms[2] if len(hms) > 2 else 0
        if ampm == "pm" and h != 12:
            h += 12
        elif ampm == "am" and h == 12:
            h = 0
        return datetime(y, m, d, h, mi, s).isoformat()
    except ValueError:
        return None


def parse_text(text, forced_order=None):
    raw, cur = [], None
    for line in text.splitlines():
        m = match(line)
        if m:
            cur = m.groupdict()
            raw.append(cur)
        elif cur is not None:          # continuation of a multi-line message
            cur["rest"] += "\n" + line
    if not raw:
        return [], None
    order = guess_order([r["date"] for r in raw], forced_order)

    msgs = []
    for r in raw:
        sender, sep, body = r["rest"].partition(": ")
        if not sep:                    # system line: "X added Y", encryption notice...
            sender, body = None, r["rest"]
        msgs.append({
            "timestamp": to_iso(r["date"], r["time"], order),
            "sender": sender.strip() if sender else None,
            "text": body.strip(),
            "is_system": sender is None,
        })
    return msgs, order


def friendly(e):
    """Turn Python's error text into something a non-programmer can act on."""
    msg = str(e)
    if "encrypted" in msg or "password" in msg:
        return "it's password-protected"
    if isinstance(e, (zipfile.BadZipFile, EOFError)) or "CRC" in msg:
        return "the file is damaged or only partly downloaded; download it again"
    if isinstance(e, NotImplementedError):
        return "it was zipped in a format Python can't open; unzip it yourself first"
    if isinstance(e, MemoryError):
        return "it's too big to open on this device"
    if isinstance(e, PermissionError):
        return "permission denied; close it in any other program and try again"
    if isinstance(e, (FileExistsError, NotADirectoryError)):
        return "there's a file with that name in the way; pick a different folder"
    if isinstance(e, OSError) and e.errno == 28:
        return "the disk is full"
    return msg or e.__class__.__name__


# ---------- finding exports ----------

def looks_like_chat(text):
    """Real exports start with a dated message line; .txt attachments sent in a chat don't."""
    for line in text.splitlines():
        if line.strip():
            return bool(match(line))
    return False


def strip_prefix(stem):
    """'WhatsApp Chat with Thabo' -> 'Thabo'"""
    low = stem.lower()
    for prefix in NAME_PREFIXES:
        if low.startswith(prefix) and len(stem) > len(prefix):
            return stem[len(prefix):].strip()
    return stem


def chat_name(path):
    """'WhatsApp Chat with Thabo.txt' -> 'Thabo'; unzipped iPhone '_chat.txt' -> its folder's name."""
    stem = path.resolve().parent.name if path.stem.lower() == "_chat" else path.stem
    return strip_prefix(stem or "chat")


def chats_in_zip(zf, zip_stem, problems, where, depth=0):
    """Every chat in a zip: one iPhone export, or a Drive/Gmail bundle of several exports."""
    found = []
    for info in zf.infolist():
        member = PurePosixPath(info.filename)
        if info.is_dir() or "__MACOSX" in member.parts or member.name.startswith("._"):
            continue
        suffix = member.suffix.lower()
        if suffix not in EXTS or (suffix == ".zip" and depth >= 2):
            continue
        try:
            data = zf.read(info)
            if suffix == ".txt":
                text = decode(data)
                if looks_like_chat(text):
                    stem = (member.parent.name or zip_stem) if member.stem.lower() == "_chat" else member.stem
                    found.append((strip_prefix(stem or "chat"), text))
            else:                      # iPhone exports inside a Drive/Gmail bundle
                with zipfile.ZipFile(io.BytesIO(data)) as inner:
                    found += chats_in_zip(inner, member.stem, problems, where + " > " + member.name, depth + 1)
        except Exception as e:         # password-protected, damaged, unusual compression...
            problems.append("{} > {}: couldn't read it, {}".format(where, member.name, friendly(e)))
    return found


def chats_in_file(path, problems):
    """(name, text) for each chat in a .txt or .zip file."""
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            return chats_in_zip(z, path.stem, problems, path.name)
    return [(chat_name(path), decode(path.read_bytes()))]


def scan_folder(folder):
    return sorted(f for f in folder.rglob("*") if f.suffix.lower() in EXTS and f.is_file())


def find_inputs(args):
    """Turn command-line inputs into (path, named_directly) pairs, plus any problems."""
    found, problems = [], []
    for arg in args:
        arg = arg.strip().strip('"')   # Windows: "C:\folder\" arrives with a stray quote
        try:
            p = Path(arg)
            if p.is_dir():
                found += [(f, False) for f in scan_folder(p)]
            elif p.is_file():
                found.append((p, True))
            elif any(ch in arg for ch in "*?["):
                hits = [Path(h) for h in sorted(glob.glob(arg, recursive=True))]
                for h in hits:
                    if h.is_dir():
                        found += [(f, False) for f in scan_folder(h)]
                    elif h.suffix.lower() in EXTS:
                        found.append((h, False))
                if not hits:
                    problems.append(arg + ": nothing matches that pattern")
            else:
                problems.append(arg + ": not found")
        except OSError as e:
            problems.append("{}: couldn't open it, {}".format(arg, friendly(e)))
    return found, problems


# ---------- WhatsApp Desktop for Mac (ChatStorage.sqlite) ----------

def _s(v):
    """Database values as text, whatever type they were stored as."""
    if isinstance(v, str):
        return v
    if isinstance(v, (bytes, bytearray)):
        return bytes(v).decode("utf-8", errors="replace")
    return "" if v is None else str(v)


def phone_from_jid(jid):
    """'27821234567@s.whatsapp.net' -> '+27821234567'"""
    jid = _s(jid)
    user = jid.split("@")[0]
    return "+" + user if jid.endswith("@s.whatsapp.net") and user.isdigit() else None


def read_chatstorage(db_path):
    """Every chat in WhatsApp's ChatStorage.sqlite. Works on a copy, never touches the original."""
    db_path = Path(str(db_path).strip().strip('"')).expanduser()
    if db_path.is_dir():               # pointed at the folder rather than the file
        db_path = db_path / "ChatStorage.sqlite"
    if not db_path.is_file():
        raise FileNotFoundError(str(db_path))
    tmp = tempfile.mkdtemp(prefix="wa_parse_")
    try:
        copy = os.path.join(tmp, "ChatStorage.sqlite")
        shutil.copyfile(str(db_path), copy)
        wal = str(db_path) + "-wal"    # recent changes not yet written into the main file
        if os.path.isfile(wal):
            shutil.copyfile(wal, copy + "-wal")
        con = sqlite3.connect(copy)
        try:
            con.text_factory = lambda b: b.decode("utf-8", errors="replace")  # never choke on bad text
            return _read_chatstorage(con)
        finally:
            con.close()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _read_chatstorage(con):
    tables = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
    if not {"ZWAMESSAGE", "ZWACHATSESSION"} <= tables:
        raise ValueError("that file isn't WhatsApp's ChatStorage.sqlite")

    push = {}
    if "ZWAPROFILEPUSHNAME" in tables:
        push = {_s(j): _s(n) for j, n in con.execute("SELECT ZJID, ZPUSHNAME FROM ZWAPROFILEPUSHNAME")}

    def display(jid, saved):
        """The name saved in your contacts (or the group's name); otherwise their full number."""
        return _s(saved).strip() or phone_from_jid(jid) or push.get(jid) or jid.split("@")[0] or "Unknown"

    sessions, names = {}, {}
    for pk, jid, saved, groupinfo in con.execute(
            "SELECT Z_PK, ZCONTACTJID, ZPARTNERNAME, ZGROUPINFO FROM ZWACHATSESSION"):
        jid = _s(jid)
        if not jid or jid in SKIP_CHATS or jid.endswith(SKIP_SUFFIXES):
            continue
        is_group = groupinfo is not None or jid.endswith("@g.us")
        sessions[pk] = (display(jid, saved), is_group)
        if not is_group:
            names[jid] = sessions[pk][0]

    member_join = "LEFT JOIN ZWAGROUPMEMBER g ON m.ZGROUPMEMBER = g.Z_PK" if "ZWAGROUPMEMBER" in tables else ""
    member_col = "g.ZMEMBERJID" if member_join else "NULL"
    rows = con.execute(
        "SELECT m.ZCHATSESSION, m.ZISFROMME, m.ZMESSAGEDATE, m.ZTEXT, m.ZMESSAGETYPE, " + member_col +
        " FROM ZWAMESSAGE m " + member_join + " ORDER BY m.ZCHATSESSION, m.ZMESSAGEDATE")

    chats = {}
    for pk, from_me, date, text, mtype, member in rows:
        if pk not in sessions:
            continue
        name, is_group = sessions[pk]
        text, member = _s(text), _s(member)
        try:
            stamp = datetime.fromtimestamp(APPLE_EPOCH + float(date)).isoformat(timespec="seconds")
        except (TypeError, ValueError, OverflowError, OSError):
            stamp = None

        is_system = False
        if from_me:
            sender = "You"
        elif is_group:                 # saved name if you have a chat with them, else their number
            sender = names.get(member) or phone_from_jid(member) or push.get(member) or member.split("@")[0] or None
        else:
            sender = name

        if mtype == 6:                 # group events: someone joined, the name changed...
            is_system, sender = True, None
            text = "Group name changed to " + text if text and not text.strip().isdigit() else "<group event>"
        elif mtype == 14:
            text = "This message was deleted"
        elif mtype in MEDIA_TYPES:
            text = text or "<{} omitted>".format(MEDIA_TYPES[mtype])
        elif not text:
            is_system, text = True, "<message type {}>".format(mtype)

        chats.setdefault(pk, []).append({"timestamp": stamp, "sender": sender, "text": text,
                                         "is_system": is_system})
    return [(sessions[pk][0], msgs) for pk, msgs in chats.items()]


# ---------- saving ----------

def safe_filename(name):
    """Folder/file name Windows, Mac and phones all accept, short enough for Windows' path limit."""
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", _s(name))[:MAX_NAME].strip(" .") or "chat"
    if name.split(".")[0].rstrip().upper() in WINDOWS_RESERVED:   # a contact saved as "Con"
        name = "_" + name
    return name


def make_folder(parent, name):
    """Create parent/name, or 'name (2)', 'name (3)'... if that name is already taken."""
    for n in range(1, 1000):
        folder = parent / (name if n == 1 else "{} ({})".format(name, n))
        try:
            folder.mkdir()
            return folder
        except FileExistsError:
            continue
    raise OSError("too many folders called " + name)


def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:  # -sig so Excel reads emojis
        w = csv.DictWriter(f, fieldnames=COLUMNS)
        w.writeheader()
        w.writerows(rows)


# A real Excel file, built by hand so nothing needs installing. Every cell is stored as
# text, so Excel can't turn "+27 82 123 4567" into a formula or 2.78E+10.
XLSX_PARTS = {
    "[Content_Types].xml":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>',
    "_rels/.rels":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>',
    "xl/workbook.xml":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Messages" sheetId="1" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels":
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '</Relationships>',
    "xl/styles.xml":                   # 0 default, 1 bold header, 2 wrapped message, 3 top-aligned
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="2"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="4">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
        '<alignment vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
        '<alignment vertical="top"/></xf>'
        '</cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>',
}


def _cell(ref, value, style):
    text = XML_BAD.sub("", _s(value))[:32767]         # Excel's limit per cell
    if not text:
        return ""
    return '<c r="{}" t="inlineStr" s="{}"><is><t xml:space="preserve">{}</t></is></c>'.format(
        ref, style, escape(text))


def write_xlsx(path, rows):
    sheet = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             '<sheetViews><sheetView workbookViewId="0">'        # header row stays put when scrolling
             '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
             '</sheetView></sheetViews><cols>']
    sheet += ['<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(i, w)
              for i, w in enumerate(WIDTHS, 1)]
    sheet.append('</cols><sheetData><row r="1">')
    sheet += [_cell(col + "1", title, 1) for col, title in zip("ABCD", COLUMNS)]
    sheet.append("</row>")
    for r, row in enumerate(rows, 2):
        sheet.append('<row r="{}">'.format(r))
        sheet += [_cell(col + str(r), row[title], 2 if title == "Message" else 3)
                  for col, title in zip("ABCD", COLUMNS)]
        sheet.append("</row>")
    sheet.append("</sheetData></worksheet>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, xml in XLSX_PARTS.items():
            z.writestr(name, xml)
        z.writestr("xl/worksheets/sheet1.xml", "".join(sheet))


def save(folder, msgs, json_too):
    """folder/<name>.xlsx + .csv (+ .json if asked), with Date | Time | Sender | Message."""
    rows = []
    for m in msgs:
        ts = m["timestamp"] or ""
        rows.append({"Date": ts[:10], "Time": ts[11:16], "Sender": m["sender"] or "", "Message": m["text"]})
    stem = str(folder / folder.name)
    if len(rows) <= EXCEL_MAX_ROWS:
        write_xlsx(stem + ".xlsx", rows)
    else:
        print("  ! {} has more messages than Excel allows; saved as .csv only".format(folder.name))
    write_csv(stem + ".csv", rows)
    if json_too:
        Path(stem + ".json").write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")


# ---------- running ----------

def describe(name, msgs, source, date_hint=True):
    stamps = [m["timestamp"] for m in msgs if m["timestamp"]]
    span = "{} to {}".format(min(stamps)[:10], max(stamps)[:10]) if stamps else "no readable dates"
    people = Counter(m["sender"] or "unknown" for m in msgs if not m["is_system"])
    top = ", ".join("{} {}".format(who, n) for who, n in people.most_common(3))
    if len(people) > 3:
        top += ", +{} more".format(len(people) - 3)
    n = len(msgs)
    line = "  OK  {}: {} message{}, {}, {} ({})".format(name, n, "" if n == 1 else "s", span, source, top)
    bad = n - len(stamps)
    if bad and date_hint:
        line += "\n      ! {} timestamps couldn't be read; try --dayfirst or --monthfirst".format(bad)
    return line


def finish(code, wait):
    """Exit; when double-clicked or tapped, wait first so the window doesn't vanish."""
    if wait and sys.stdin is not None and sys.stdin.isatty():
        try:
            input("\nPress Enter to close...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(code)


def main():
    try:
        sys.stdout.reconfigure(errors="replace")  # don't crash on emoji chat names in odd consoles
    except AttributeError:
        pass

    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="*", metavar="INPUT",
                    help="a .txt or .zip export, a folder of exports, or a wildcard pattern")
    ap.add_argument("--mac-app", nargs="?", const=str(MAC_DB), metavar="FILE",
                    help="read WhatsApp Desktop's database on this Mac (or a copied ChatStorage.sqlite FILE)")
    ap.add_argument("--out", help="folder for the results (default: parsed)")
    ap.add_argument("--json", action="store_true", help="also save .json files (for programs)")
    g = ap.add_mutually_exclusive_group()
    g.add_argument("--dayfirst", action="store_const", const="dmy", dest="order",
                   help="read ambiguous dates as DD/MM")
    g.add_argument("--monthfirst", action="store_const", const="mdy", dest="order",
                   help="read ambiguous dates as MM/DD")
    a = ap.parse_args()

    # No arguments (double-click / phone "Run" button): use the folders next to this script
    simple = not a.inputs and a.mac_app is None
    try:
        here = Path(__file__).resolve().parent
    except NameError:                  # some phone Python apps don't set __file__
        here = Path.cwd()
    if simple:
        exports = here / "exports"
        if not exports.is_dir():
            try:
                exports.mkdir(parents=True)
                print("Created the folder {}\nPut your exported chats in it, then run this again.".format(exports))
                finish(0, simple)
            except OSError as e:
                print("Couldn't create {} ({}).\nMake a folder called 'exports' next to wa_parse.py "
                      "yourself, put your exported chats in it, then run this again.".format(exports, e))
                finish(1, simple)
        a.inputs = [str(exports)]
        out = Path(a.out) if a.out else here / "parsed"
    else:
        out = Path(a.out or "parsed")

    candidates, problems = find_inputs(a.inputs)
    seen, used, results, skipped = {}, set(), [], 0

    def add_result(name, msgs, source, date_hint=True):
        """Pick a folder name nobody else in this run has (Windows ignores case, Macs ignore accents)."""
        base = safe_filename(name)
        key, n = base, 2
        while unicodedata.normalize("NFC", key).casefold() in used:
            key, n = "{} ({})".format(base[:MAX_NAME - 6], n), n + 1
        used.add(unicodedata.normalize("NFC", key).casefold())
        shown = name if key == base else key
        results.append((key, msgs))
        print(describe(shown, msgs, source, date_hint))
        return shown

    # 1) exported chats: one bad file never stops the rest
    for path, named_directly in candidates:
        before = len(problems)
        try:
            chats = chats_in_file(path, problems)
        except Exception as e:
            problems.append("{}: couldn't read it, {}".format(path, friendly(e)))
            continue

        found_chat = False
        for name, text in chats:
            try:
                digest = hashlib.sha256(text.encode("utf-8", errors="replace")).hexdigest()
                if digest in seen:
                    where = "{} in {}".format(name, path.name) if path.suffix.lower() == ".zip" else path.name
                    print("  --  {}: same chat as '{}', skipped".format(where, seen[digest]))
                    found_chat = True
                    continue
                msgs, order = parse_text(text, a.order)
                if not msgs:
                    continue
                found_chat = True
                seen[digest] = add_result(name, msgs, "dates read " + ORDER_LABELS[order])
            except Exception as e:
                problems.append("{} ({}): couldn't read it, {}".format(name, path.name, friendly(e)))

        if not found_chat and len(problems) == before:
            if named_directly:
                problems.append("{}: no messages matched. Paste its first 3 lines "
                                "(names changed) so the pattern can be adjusted".format(path))
            else:
                skipped += 1           # e.g. a README or other stray file in the folder

    # 2) WhatsApp Desktop for Mac
    if a.mac_app is not None:
        try:
            db_chats = read_chatstorage(a.mac_app)
        except FileNotFoundError:
            problems.append("{}: not found. --mac-app needs the WhatsApp app for Mac installed and signed in "
                            "on this Mac (WhatsApp Web in a browser has no database)".format(a.mac_app))
        except PermissionError:
            problems.append("macOS blocked access to WhatsApp's data. Allow Terminal when macOS asks, or turn it on "
                            "in System Settings > Privacy & Security > Full Disk Access, then run this again")
        except ValueError as e:
            problems.append("{}: {}".format(a.mac_app, e))
        except sqlite3.DatabaseError as e:
            problems.append("{}: couldn't read it ({}). If WhatsApp was open, quit it (Cmd+Q) "
                            "and try again".format(a.mac_app, e))
        except Exception as e:
            problems.append("{}: couldn't read it, {}".format(a.mac_app, friendly(e)))
        else:
            for name, msgs in db_chats:
                add_result(name, msgs, "from WhatsApp for Mac", date_hint=False)
            if not db_chats:
                problems.append("the WhatsApp for Mac database has no chats in it yet")

    # 3) save: parsed/2026-11-10 09h42/Thabo Mokoena/Thabo Mokoena.xlsx
    saved = 0
    if results:
        try:
            out.mkdir(parents=True, exist_ok=True)
            run = make_folder(out, datetime.now().strftime("%Y-%m-%d %Hh%M"))
        except OSError as e:
            problems.append("couldn't create the output folder {}, {}".format(out, friendly(e)))
            run = None
        if run is not None:
            for key, msgs in results:
                try:
                    save(make_folder(run, key), msgs, a.json)
                    saved += 1
                except OSError as e:   # e.g. disk full, or a file open in another program
                    problems.append("{}: couldn't save it, {}".format(key, friendly(e)))
            print("\nSaved {} chat(s) to {}".format(saved, run.resolve()))

    if skipped:
        print("Skipped {} file(s) that weren't WhatsApp exports".format(skipped))
    if problems:
        print("\nProblems:")
        for p in problems:
            print("  !!  " + p)
    if not results:
        print("\nNo chats were found.")
        if simple:
            print("Put your exported chats in {}, then run this again.".format(a.inputs[0]))
    finish(1 if problems or not saved else 0, simple)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
        sys.exit(130)
    except Exception:                  # show the error instead of a window that just closes
        traceback.print_exc()
        print("\nSomething unexpected went wrong. Copy the error above and send it over.")
        finish(1, len(sys.argv) == 1)
