import os
import re
import gspread
import pyperclip
import json
import datetime

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, Input, ListView, ListItem, Button, TextArea
from textual.containers import Vertical, Horizontal, VerticalScroll
from textual.screen import Screen


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
MEMORY_FILE = "sheets_memory.json"
MAX_RANGE = 30  # 🔒 Hard limit for any range command


# =============================
# SIMPLE MEMORY SYSTEM
# =============================

def load_memory():
    if not os.path.exists(MEMORY_FILE):
        return []
    with open(MEMORY_FILE, "r") as f:
        return json.load(f)


def save_memory(memory):
    with open(MEMORY_FILE, "w") as f:
        json.dump(memory, f, indent=2)


def add_or_update_memory(url, title):
    memory = load_memory()
    now = datetime.datetime.now().isoformat(timespec="seconds")

    for item in memory:
        if item["url"] == url:
            item["last_accessed"] = now
            save_memory(memory)
            return

    new_id = max([m["id"] for m in memory], default=0) + 1

    memory.append({
        "id": new_id,
        "url": url,
        "title": title,
        "last_accessed": now
    })

    save_memory(memory)


def select_sheet_from_memory():
    memory = load_memory()

    if memory:
        print("\nSaved Sheets:\n")
        for item in memory:
            print(
                f"[{item['id']}] {item['title']}\n"
                f"     {item['url']}\n"
                f"     Last accessed: {item['last_accessed']}\n"
            )
    else:
        print("No saved sheets.\n")

    choice = input("Enter sheet ID or paste new Google Sheets URL: ").strip()

    if choice.isdigit():
        selected_id = int(choice)
        for item in memory:
            if item["id"] == selected_id:
                return item["url"]
        print("Invalid ID.")
        return select_sheet_from_memory()

    return choice


# =============================
# GOOGLE AUTH
# =============================

def extract_sheet_id(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None


def get_credentials():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    try:
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                raise Exception("Re-auth required")

    except Exception:
        if os.path.exists("token.json"):
            os.remove("token.json")

        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def connect_sheet(sheet_url):
    sheet_id = extract_sheet_id(sheet_url)
    creds = get_credentials()
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id)


# =============================
# SESSION
# =============================

class SheetSession:
    def __init__(self, spreadsheet):
        self.spreadsheet = spreadsheet
        self.sheet = spreadsheet.sheet1
        self.data = []

    def load(self):
        self.data = self.sheet.get_all_values()

    def row_count(self):
        return len(self.data)

    def get_cell(self, idx, col_index):
        if len(self.data[idx]) > col_index:
            return self.data[idx][col_index]
        return ""

    def bulk_update_contiguous_i(self, start, end, new_values):
        if not new_values:
            return "No changes."

        values_2d = [[v] for v in new_values]

        self.sheet.batch_update([{
            "range": f"I{start+1}:I{end+1}",
            "values": values_2d
        }])

        for offset, value in enumerate(new_values):
            self.data[start + offset][8] = value

        return f"{len(new_values)} rows saved successfully."


# =============================
# RANGE VALIDATION MIXIN
# =============================

class RangeValidator:

    def validate_range(self, start, end):
        if start < 0 or end >= self.session.row_count() or end < start:
            return "Invalid range."

        if end - start + 1 > MAX_RANGE:
            return f"Maximum {MAX_RANGE} rows allowed."

        return None


# =============================
# BATCH EDIT SCREEN
# =============================

class BatchEditScreen(Screen):

    def __init__(self, session, start, end):
        super().__init__()
        self.session = session
        self.start = start
        self.end = end
        self.current = start

        self.edited_values = [
            self.session.get_cell(i, 8)
            for i in range(start, end + 1)
        ]

    def compose(self) -> ComposeResult:
        self.content = Static()
        self.input = Input(placeholder="New I value (blank = keep)")
        yield Vertical(
            Static(f"Batch Edit Mode — {self.end - self.start + 1} rows (ESC = cancel)\n"),
            self.content,
            self.input
        )

    def on_mount(self):
        self.load_row()

    def load_row(self):
        if self.current > self.end:
            msg = self.session.bulk_update_contiguous_i(
                self.start,
                self.end,
                self.edited_values
            )
            self.app.notify(msg)
            self.app.pop_screen()
            self.app.refresh_main_list()
            return

        f = self.session.get_cell(self.current, 5)
        g = self.session.get_cell(self.current, 6)
        h = self.session.get_cell(self.current, 7)
        i = self.session.get_cell(self.current, 8)

        self.content.update(
            f"Row {self.current+1}\n\n"
            f"F: {f}\nG: {g}\nH: {h}\nI: {i}\n"
        )

        self.input.value = ""
        self.input.focus()

    def on_input_submitted(self, event: Input.Submitted):
        event.stop()

        new_value = event.value.strip()

        if new_value:
            offset = self.current - self.start
            self.edited_values[offset] = new_value

        self.current += 1
        self.load_row()

    def on_key(self, event):
        if event.key == "escape":
            self.app.notify("Batch cancelled.")
            self.app.pop_screen()


# =============================
# BATCH PASTE SCREEN
# =============================

class BatchPasteScreen(Screen):

    BINDINGS = [
        ("v", "paste", "Paste"),
        ("s", "save", "Save"),
        ("c", "copy", "Copy"),
        ("escape", "cancel", "Cancel"),
    ]

    def __init__(self, session, start, end, value_data=None):
        super().__init__()
        self.session = session
        self.start = start
        self.end = end
        self.total_lines = end - start + 1
        self.value_data = value_data
        self.row_widgets = []

    def compose(self) -> ComposeResult:
        yield Static(
            f"Batch Paste Mode — {self.total_lines} rows selected "
            "(V = Paste | S = Save | C = Copy | ESC Cancel)\n"
        )
        
        self.scroll = VerticalScroll()
        with self.scroll:
            for idx in range(self.start, self.end + 1):
                f = self.session.get_cell(idx, 5)
                g = self.session.get_cell(idx, 6)
                h = self.session.get_cell(idx, 7)
                i = self.session.get_cell(idx, 8)

                row_widget = Static(
                    f"Row {idx+1}\nF: {f}\nG: {g}\nH: {h}\nI: {i}\n"
                )

                self.row_widgets.append(row_widget)
                yield row_widget

        self.scroll.styles.height = "20" 
        yield self.scroll

        yield Horizontal(
            Button("Paste from Clipboard", id="paste"),
            Button("Save", id="save"),
            Button("Copy", id="copy"),
        )

        self.text_area = TextArea()
        self.text_area.display = False
        yield self.text_area

    def update_preview(self, new_lines):
        for offset, line in enumerate(new_lines):
            idx = self.start + offset
            f = self.session.get_cell(idx, 5)
            g = self.session.get_cell(idx, 6)
            h = self.session.get_cell(idx, 7)

            self.row_widgets[offset].update(
                f"Row {idx+1}\nF: {f}\nG: {g}\nH: {h}\nI: {line}\n"
            )

    def action_paste(self):
        expected = self.total_lines
        clipboard_data = pyperclip.paste()
        lines = clipboard_data.splitlines()

        self.text_area.text = clipboard_data

        if len(lines) != expected:
            self.app.notify(f"Expected {expected} lines, got {len(lines)}.")
            return

        self.update_preview(lines)

    def action_save(self):
        lines = self.text_area.text.splitlines()

        if len(lines) != self.total_lines:
            self.app.notify(f"Expected {self.total_lines} lines.")
            return

        msg = self.session.bulk_update_contiguous_i(
            self.start,
            self.end,
            lines
        )

        self.app.notify(msg)
        self.app.pop_screen()
        self.app.refresh_main_list()

    def action_copy(self):
        pyperclip.copy(self.value_data)

    def action_cancel(self):
        self.app.notify("Batch cancelled.")
        self.app.pop_screen()

    def on_button_pressed(self, event: Button.Pressed):
        if event.button.id == "paste":
            self.action_paste()
        elif event.button.id == "save":
            self.action_save()
        elif event.button.id == "copy":
            self.action_copy()


# =============================
# MAIN APP
# =============================

class SheetApp(App, RangeValidator):

    def __init__(self, session):
        super().__init__()
        self.session = session
        self.list_view = ListView()

    def compose(self) -> ComposeResult:
        yield Header()
        yield self.list_view
        yield Input(placeholder="command")
        yield Footer()

    def on_mount(self):
        self.populate()

    def populate(self):
        self.list_view.clear()
        for idx in range(self.session.row_count()):
            g = self.session.get_cell(idx, 6)
            i = self.session.get_cell(idx, 8)
            self.list_view.append(
                ListItem(Static(f"{idx+1}. {g}\n   {i}"))
            )

    def refresh_main_list(self):
        self.populate()

    def parse_row_arg(self, raw_value):
        raw_value = raw_value.strip()
        if not raw_value.isdigit():
            return None
        return int(raw_value) - 1

    def on_input_submitted(self, event: Input.Submitted):
        cmd = event.value.strip()
        event.input.value = ""
        if not cmd:
            return

        parts = cmd.split()
        command = parts[0].lower()

        if command in ("quit", "q"):
            self.exit()

        elif command in ("reload", "r"):
            self.session.load()
            self.populate()
            self.notify("Reloaded.")

        elif command in ("go", "g") and len(parts) == 2 and parts[1].isdigit():
            row = int(parts[1]) - 1
            if 0 <= row < self.session.row_count():
                self.list_view.index = row
            else:
                self.notify("Out of range")

        elif command in ("copy", "c") and len(parts) in (2, 3):
            start = self.parse_row_arg(parts[1])
            end = start if len(parts) == 2 else self.parse_row_arg(parts[2])

            if start is None or end is None:
                self.notify("Invalid row number. Use whole numbers, e.g. 'batch 10 20'.")
                return

            error = self.validate_range(start, end)
            if error:
                self.notify(error)
                return

            values = [
                self.session.get_cell(i, 5)
                for i in range(start, end + 1)
            ]
            pyperclip.copy("\n".join(values))
            self.notify(f"Copied {end-start+1} row(s).")

        elif command in ("edit", "e") and len(parts) in (2, 3):
            start = self.parse_row_arg(parts[1])
            end = start if len(parts) == 2 else self.parse_row_arg(parts[2])

            if start is None or end is None:
                self.notify("Invalid row number. Use whole numbers, e.g. 'batch 10 20'.")
                return

            error = self.validate_range(start, end)
            if error:
                self.notify(error)
                return

            self.push_screen(BatchEditScreen(self.session, start, end))

        elif command in ("batch", "b") and len(parts) == 3:
            start = self.parse_row_arg(parts[1])
            end = self.parse_row_arg(parts[2])

            if start is None or end is None:
                self.notify("Invalid row number. Use whole numbers, e.g. 'batch 10 20'.")
                return

            error = self.validate_range(start, end)
            if error:
                self.notify(error)
                return

            values = [
                self.session.get_cell(i, 5)
                for i in range(start, end + 1)
            ]
            value_data = "\n".join(values)
            pyperclip.copy(value_data)
            self.push_screen(BatchPasteScreen(self.session, start, end, value_data))

        else:
            self.notify("Unknown command")


# =============================
# ENTRY POINT
# =============================

def main():
    url = select_sheet_from_memory()
    spreadsheet = connect_sheet(url)
    add_or_update_memory(url, spreadsheet.title)

    session = SheetSession(spreadsheet)
    session.load()

    app = SheetApp(session)
    app.run()


if __name__ == "__main__":
    main()