import os
import re
import gspread

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from textual.app import App, ComposeResult
from textual.widgets import Header, Footer, Static, Input, ListView, ListItem
from textual.containers import Vertical
from textual.screen import Screen


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


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

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_console()

        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds


def connect_sheet(sheet_url):
    sheet_id = extract_sheet_id(sheet_url)
    creds = get_credentials()
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id).sheet1


# =============================
# SESSION
# =============================

class SheetSession:
    def __init__(self, sheet):
        self.sheet = sheet
        self.data = []

    def load(self):
        self.data = self.sheet.get_all_values()

    def row_count(self):
        return len(self.data)

    def get_cell(self, idx, col_index):
        if len(self.data[idx]) > col_index:
            return self.data[idx][col_index]
        return ""

    def bulk_update_i(self, updates: dict):
        """
        updates = { row_index: new_value }
        """
        if not updates:
            return "No changes."

        requests = []
        for idx, value in updates.items():
            requests.append({
                "range": f"I{idx+1}",
                "values": [[value]]
            })

        self.sheet.batch_update(requests)

        # Update local cache
        for idx, value in updates.items():
            self.data[idx][8] = value

        return f"{len(updates)} rows saved."


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
        self.pending_updates = {}   # store changes here

    def compose(self) -> ComposeResult:
        self.content = Static()
        self.input = Input(placeholder="New I value (blank = keep)")
        yield Vertical(
            Static("Batch Edit Mode (ESC to cancel)\n"),
            self.content,
            self.input
        )

    def on_mount(self):
        self.load_row()

    def load_row(self):
        if self.current > self.end:
            # Batch finished → send one API call
            msg = self.session.bulk_update_i(self.pending_updates)
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
            f"F: {f}\n"
            f"G: {g}\n"
            f"H: {h}\n"
            f"I: {i}\n"
        )

        self.input.value = ""
        self.input.focus()

    def on_input_submitted(self, event: Input.Submitted):
        new_value = event.value.strip()

        if new_value:
            self.pending_updates[self.current] = new_value

        self.current += 1
        self.load_row()

    def on_key(self, event):
        if event.key == "escape":
            self.app.notify("Batch cancelled. No changes saved.")
            self.app.pop_screen()
            self.app.refresh_main_list()


# =============================
# HELP SCREEN
# =============================

class HelpScreen(Screen):

    def compose(self) -> ComposeResult:
        yield Static("""
COMMANDS

:help
    Show this screen.

:edit <start> <end>
    Batch edit rows (auto-save each row).

:save
    (Not needed anymore — auto-save enabled.)

:reload
    Reload data from Google Sheets.

:go <row>
    Jump to a row.

:q
    Quit application.

Navigation:
Arrow keys to move.
ESC exits batch edit.
Press any key to return.
""")

    def on_key(self, event):
        self.app.pop_screen()


# =============================
# MAIN APP
# =============================

class SheetApp(App):

    def __init__(self, session):
        super().__init__()
        self.session = session
        self.list_view = ListView()

    def compose(self) -> ComposeResult:
        yield Header()
        yield self.list_view
        yield Input(placeholder=":command")
        yield Footer()

    def on_mount(self):
        self.populate()

    def populate(self):
        self.list_view.clear()

        for idx in range(self.session.row_count()):
            g = self.session.get_cell(idx, 6)
            i = self.session.get_cell(idx, 8)
            text = f"{idx+1}. {g}\n   {i}"
            self.list_view.append(ListItem(Static(text)))

    def refresh_main_list(self):
        self.populate()

    def on_input_submitted(self, event: Input.Submitted):
        cmd = event.value.strip()
        event.input.value = ""

        if cmd == ":q":
            self.exit()

        elif cmd == ":help":
            self.push_screen(HelpScreen())

        elif cmd == ":reload":
            self.session.load()
            self.populate()
            self.notify("Reloaded.")

        elif cmd.startswith(":go"):
            parts = cmd.split()
            if len(parts) == 2 and parts[1].isdigit():
                row = int(parts[1]) - 1
                if 0 <= row < self.session.row_count():
                    self.list_view.index = row
                else:
                    self.notify("Out of range")

        elif cmd.startswith(":edit"):
            parts = cmd.split()
            if len(parts) == 3 and parts[1].isdigit() and parts[2].isdigit():
                start = int(parts[1]) - 1
                end = int(parts[2]) - 1

                if 0 <= start <= end < self.session.row_count():
                    self.push_screen(BatchEditScreen(self.session, start, end))
                else:
                    self.notify("Invalid range")


# =============================

def main():
    url = input("Enter Google Sheets URL: ")
    sheet = connect_sheet(url)

    session = SheetSession(sheet)
    session.load()

    app = SheetApp(session)
    app.run()


if __name__ == "__main__":
    main()