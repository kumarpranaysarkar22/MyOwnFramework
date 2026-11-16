import ctypes
from io import BytesIO
from importlib import import_module
from pathlib import Path
import time
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog
from ctypes import wintypes
from zipfile import BadZipFile

try:
    Image = import_module("PIL.Image")
    ImageGrab = import_module("PIL.ImageGrab")
except ImportError:
    Image = None
    ImageGrab = None

try:
    docx = import_module("docx")
    Document = docx.Document
    Inches = import_module("docx.shared").Inches
    PackageNotFoundError = import_module("docx.opc.exceptions").PackageNotFoundError
except ImportError:
    Document = None
    Inches = None
    PackageNotFoundError = None


class ScreenshotApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Screenshot Capture")
        self.root.resizable(False, False)
        self.document_path = None
        self.name_screenshots = None
        self.has_used_application = False
        self.has_captured_screenshot = False
        documents_directory = Path.home() / "Documents"
        self.last_document_directory = (
            documents_directory if documents_directory.is_dir() else Path.home()
        )
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)
        popup_background = "#F3F4F6"
        self.root.configure(bg=popup_background)

        buttons = tk.Frame(root, padx=10, pady=8, bg=popup_background)
        buttons.pack(anchor=tk.CENTER)

        self.button_font = tkfont.nametofont("TkDefaultFont").copy()
        self.button_font.configure(
            size=self.button_font.cget("size") + 2,
            weight="bold",
        )
        button_style = {
            "width": 7,
            "font": self.button_font,
            "relief": tk.FLAT,
            "borderwidth": 0,
            "highlightthickness": 0,
            "padx": 11,
            "pady": 6,
            "cursor": "hand2",
            "background": "#E5E7EB",
            "foreground": "#1F2937",
            "activebackground": "#D1D5DB",
            "activeforeground": "#1F2937",
        }
        open_button_style = {
            **button_style,
            "background": "#DBEAFE",
            "foreground": "#1E40AF",
            "activebackground": "#BFDBFE",
            "activeforeground": "#1E40AF",
        }
        self.open_button = tk.Button(
            buttons,
            text="Open",
            command=self.choose_word_document,
            **open_button_style,
        )
        self.open_button.pack(side=tk.LEFT, padx=4)
        self.capture_button = tk.Button(
            buttons,
            text="Click",
            command=self.capture_screenshot,
            state=tk.DISABLED,
            **button_style,
        )
        self.capture_button.pack(side=tk.LEFT, padx=4)
        self.close_button = tk.Button(
            buttons,
            text="Close",
            command=self.close_application,
            **button_style,
        )
        self.close_button.pack(side=tk.LEFT, padx=4)
        self.set_document_controls(False)

        self.root.update_idletasks()
        width = max(self.root.winfo_reqwidth(), 275)
        height = self.root.winfo_reqheight()
        work_area = wintypes.RECT()
        system_parameters_info = ctypes.WinDLL(
            "user32", use_last_error=True
        ).SystemParametersInfoW
        system_parameters_info.argtypes = (
            wintypes.UINT,
            wintypes.UINT,
            ctypes.POINTER(wintypes.RECT),
            wintypes.UINT,
        )
        system_parameters_info.restype = wintypes.BOOL
        if not system_parameters_info(
            0x0030, 0, ctypes.byref(work_area), 0
        ):
            raise ctypes.WinError(ctypes.get_last_error())

        x = max(work_area.left, work_area.right - width - 10)
        y = max(work_area.top, work_area.bottom - height - 30)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def choose_word_document(self) -> None:
        if Document is None:
            self.show_message(
                "Missing dependency",
                "Install python-docx to save screenshots as Word documents.",
            )
            return

        path = filedialog.asksaveasfilename(
            title="Open an existing document or create a new one",
            initialdir=str(self.last_document_directory),
            initialfile="screenshots.docx",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")],
            confirmoverwrite=False,
        )
        if not path:
            return

        try:
            if Path(path).exists():
                Document(path)
            else:
                document_content = BytesIO()
                try:
                    Document().save(document_content)
                    try:
                        with open(path, "xb") as new_file:
                            new_file.write(document_content.getvalue())
                    except FileExistsError:
                        Document(path)
                finally:
                    document_content.close()
        except (OSError, ValueError, BadZipFile, PackageNotFoundError) as error:
            self.show_message("Document error", str(error))
            return

        self.last_document_directory = Path(path).parent
        self.activate_document(path)

    def activate_document(self, path: str) -> None:
        self.document_path = path
        self.has_captured_screenshot = False
        self.name_screenshots = self.ask_yes_no(
            "Name screenshots?",
            "Would you like to name screenshots in this document? Yes asks for a "
            "name on every capture; No saves captures without names.",
        )
        self.set_document_controls(True)
        self.has_used_application = True

    def set_document_controls(self, document_selected: bool) -> None:
        if document_selected:
            self.open_button.config(
                state=tk.DISABLED,
                background="#E5E7EB",
                foreground="#9CA3AF",
                activebackground="#E5E7EB",
                activeforeground="#9CA3AF",
                disabledforeground="#9CA3AF",
            )
            self.capture_button.config(
                state=tk.NORMAL,
                background="#BFDBFE",
                foreground="#1E3A8A",
                activebackground="#93C5FD",
                activeforeground="#1E3A8A",
            )
        else:
            self.open_button.config(
                state=tk.NORMAL,
                background="#DBEAFE",
                foreground="#1E40AF",
                activebackground="#BFDBFE",
                activeforeground="#1E40AF",
            )
            self.capture_button.config(
                state=tk.DISABLED,
                background="#E5E7EB",
                foreground="#9CA3AF",
                activebackground="#E5E7EB",
                activeforeground="#9CA3AF",
                disabledforeground="#9CA3AF",
            )

    def show_styled_dialog(
        self,
        title: str,
        message: str,
        actions: tuple[tuple[str, object], ...],
        input_field: bool = False,
    ):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.configure(bg="#F3F4F6")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        content = tk.Frame(dialog, bg="#F3F4F6", padx=22, pady=18)
        content.pack(fill=tk.BOTH, expand=True)
        tk.Label(
            content,
            text=message,
            justify=tk.LEFT,
            anchor="w",
            wraplength=440,
            bg="#F3F4F6",
            fg="#1F2937",
        ).pack(fill=tk.X)

        entry = None
        if input_field:
            entry = tk.Entry(content, width=40)
            entry.pack(fill=tk.X, pady=(12, 0))

        result = {"value": None}

        def finish(value):
            result["value"] = entry.get() if value == "entry" else value
            dialog.destroy()

        button_row = tk.Frame(content, bg="#F3F4F6")
        button_row.pack(fill=tk.X, pady=(18, 0))
        for label, value in reversed(actions):
            tk.Button(
                button_row,
                text=label,
                command=lambda action_value=value: finish(action_value),
                font=self.button_font,
                relief=tk.FLAT,
                borderwidth=0,
                highlightthickness=0,
                padx=14,
                pady=7,
                cursor="hand2",
                background="#DBEAFE",
                foreground="#1E3A8A",
                activebackground="#BFDBFE",
                activeforeground="#1E3A8A",
            ).pack(side=tk.RIGHT, padx=(8, 0))

        enter_action = actions[0][1]
        escape_action = actions[1][1] if len(actions) > 1 else None
        dialog.update_idletasks()
        width = max(dialog.winfo_reqwidth(), 320)
        height = dialog.winfo_reqheight()
        x = max(0, (dialog.winfo_screenwidth() - width) // 2)
        y = max(0, (dialog.winfo_screenheight() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.protocol("WM_DELETE_WINDOW", lambda: finish(None))
        dialog.bind("<Escape>", lambda _event: finish(escape_action))
        dialog.bind("<Return>", lambda _event: finish(enter_action))
        dialog.bind("<KP_Enter>", lambda _event: finish(enter_action))
        if entry is not None:
            entry.bind("<Escape>", lambda _event: finish(escape_action))
            entry.bind("<Return>", lambda _event: finish(enter_action))
            entry.bind("<KP_Enter>", lambda _event: finish(enter_action))
        dialog.grab_set()
        dialog.focus_force()
        if entry is not None:
            entry.focus_set()
        dialog.wait_window()
        return result["value"]

    def show_message(self, title: str, message: str) -> None:
        self.show_styled_dialog(title, message, (("OK", True),))

    def ask_yes_no(self, title: str, message: str) -> bool:
        return self.show_styled_dialog(
            title,
            message,
            (("Yes", True), ("No", False)),
        ) is True

    def ask_string(self, title: str, message: str):
        return self.show_styled_dialog(
            title,
            message,
            (("OK", "entry"), ("Cancel", None)),
            input_field=True,
        )

    def close_application(self) -> None:
        if self.document_path is None:
            if not self.has_used_application:
                self.show_message(
                    "Application not used",
                    "No Word document was opened and no screenshot was saved.",
                )
            self.root.destroy()
            return

        if not self.ask_yes_no(
            "Finish screenshot capture?",
            "Are you done capturing screenshots?",
        ):
            return

        captured_screenshot = self.has_captured_screenshot
        self.close_current_document()
        if not captured_screenshot:
            self.show_message(
                "No screenshot captured",
                "You have not captured any screenshot on the file created.",
            )

    def close_current_document(self) -> None:
        self.document_path = None
        self.name_screenshots = None
        self.has_captured_screenshot = False
        self.set_document_controls(False)

    def capture_screenshot(self) -> None:
        if Image is None or ImageGrab is None:
            self.show_message(
                "Missing dependency",
                "Install Pillow to capture screenshots.",
            )
            return

        if self.document_path is None:
            self.show_message(
                "Choose a document",
                "Choose a Word document before capturing screenshots.",
            )
            return

        screenshot_name = ""
        if self.name_screenshots:
            entered_name = self.ask_string(
                "Screenshot name",
                "Enter a name for this screenshot, or leave it blank:",
            )
            if entered_name is not None:
                screenshot_name = entered_name.strip()
            time.sleep(0.5)

        heading_saved = bool(screenshot_name)
        if heading_saved and not self.append_screenshot_name(screenshot_name):
            return

        try:
            screenshot = self.grab_screen_without_window()
        except OSError as error:
            self.show_message("Screenshot failed", str(error))
            return

        screenshot_bytes = BytesIO()
        try:
            rgb_screenshot = screenshot.convert("RGB")
            try:
                rgb_screenshot.save(screenshot_bytes, format="PNG")
                screenshot_data = screenshot_bytes.getvalue()
            finally:
                rgb_screenshot.close()
        except OSError as error:
            self.show_message("Screenshot failed", str(error))
            return
        finally:
            screenshot.close()
            screenshot_bytes.close()

        if self.append_screenshot(screenshot_data, heading_saved):
            self.has_captured_screenshot = True

    def grab_screen_without_window(self):
        try:
            self.root.withdraw()
            self.root.update()
            time.sleep(0.3)
            return ImageGrab.grab()
        finally:
            self.root.deiconify()
            self.root.lift()
            self.root.update()

    def append_screenshot_name(self, screenshot_name: str) -> bool:
        if Document is None or self.document_path is None:
            return False

        try:
            document = Document(self.document_path)
            if document.paragraphs or document.tables:
                document.add_page_break()
            document.add_heading(screenshot_name, level=2)
            document.save(self.document_path)
        except (OSError, ValueError, BadZipFile, PackageNotFoundError) as error:
            self.show_message("Save failed", str(error))
            return False

        return True

    def append_screenshot(self, screenshot_data: bytes, heading_saved: bool) -> bool:
        if Document is None or Inches is None or self.document_path is None:
            return False

        try:
            document = Document(self.document_path)
            if not heading_saved and (document.paragraphs or document.tables):
                document.add_page_break()

            document.add_picture(BytesIO(screenshot_data), width=Inches(6.5))
            document.save(self.document_path)
        except (OSError, ValueError, BadZipFile, PackageNotFoundError) as error:
            self.show_message("Save failed", str(error))
            return False

        return True

def main() -> None:
    root = tk.Tk()
    ScreenshotApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()