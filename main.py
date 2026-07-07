#!/usr/bin/env python3
"""Milestone Maker -- GUI front-end for the schedule-deck engine.

A small window (no file dialog on launch): a path box + 찾아보기 to pick the
input .xlsx/.xlsm, a 생성 button that asks where to save and then builds the
deck from the EMBEDDED template. Program version and the distributor stamp
(VoE) are shown in the window. After a successful build the save folder opens.

Run with no args -> GUI. Run with --input/--output -> headless CLI (for tests
and verify.py); the GUI path and the CLI path call the same engine2.generate.
"""
import argparse
import datetime
import os
import shutil
import subprocess
import sys
import zipfile
import traceback

import engine2
from engine import resource_path, app_dir

APP_NAME = "Milestone Maker"
VERSION = "v1.0"                                  # bump to v1.1.. (small) / v2.0 (major)
VOE = ""
TEMPLATE_RESOURCE = "output.pptx"                 # embedded via --add-data
TEMPLATE_INPUT_RESOURCE = "template_input.xlsm"   # embedded blank input form (.xlsm + macro/button)
ICON_RESOURCE = "app.ico"                         # embedded window/taskbar icon


def _template_path():
    """Embedded PPT template: bundle root when frozen, assets/ when run from source."""
    if getattr(sys, "frozen", False):
        return resource_path(TEMPLATE_RESOURCE)
    return resource_path(os.path.join("assets", TEMPLATE_RESOURCE))


def _input_template_path():
    """Embedded blank INPUT form (.xlsm with macro + button) the user can save out."""
    if getattr(sys, "frozen", False):
        return resource_path(TEMPLATE_INPUT_RESOURCE)
    return resource_path(os.path.join("assets", TEMPLATE_INPUT_RESOURCE))


def _icon_path():
    """Embedded .ico: bundle root when frozen, assets/ when run from source."""
    if getattr(sys, "frozen", False):
        return resource_path(ICON_RESOURCE)
    return resource_path(os.path.join("assets", ICON_RESOURCE))


def _log_error(exc):
    """Append a timestamped traceback to error.log beside the exe; return path."""
    path = os.path.join(app_dir(), "error.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"\n=== {datetime.datetime.now().isoformat()} ===\n")
        f.write("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    return path


def _open_folder(file_path):
    """Open the folder that contains file_path (Explorer/Finder/xdg)."""
    folder = os.path.dirname(os.path.abspath(file_path))
    try:
        if sys.platform.startswith("win"):
            os.startfile(folder)  # noqa: BLE001 (Windows only)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])
    except Exception:
        pass


def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.title(f"{APP_NAME}({VERSION})")
    root.geometry("560x270")
    root.resizable(False, False)
    try:
        root.iconbitmap(_icon_path())                # window + taskbar icon
    except Exception:
        pass                                         # never block the GUI over an icon

    bg_var = tk.StringVar(value="none")              # PPT 배경: 흰 배경 고정

    # distributor stamp, top-right, small/grey

    tk.Label(root, text=APP_NAME, font=("Segoe UI", 16, "bold")).place(x=18, y=12)
    tk.Label(root, text=VERSION, fg="#666666", font=("Segoe UI", 9)).place(x=20, y=46)

    tk.Label(root, text="입력 파일 (.xlsx / .xlsm)", font=("Segoe UI", 9)).place(x=18, y=84)
    path_var = tk.StringVar()
    entry = tk.Entry(root, textvariable=path_var, width=58)
    entry.place(x=18, y=106)

    def browse():
        p = filedialog.askopenfilename(
            parent=root,
            title="입력 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if p:
            path_var.set(p)

    tk.Button(root, text="찾아보기", width=10, command=browse).place(x=458, y=103)

    # Re-entrancy guard: asksaveasfilename / showinfo each spin a NESTED modal
    # event loop. While the (synchronous) build runs the window looks frozen, so
    # impatient extra clicks on the buttons get queued and then delivered inside
    # those nested loops -> generate()/save_input_template() would run again and
    # stack another "save location" dialog. The flag + disabled button stop that.
    busy = {"on": False}

    def generate():
        if busy["on"]:
            return
        busy["on"] = True
        gen_btn.config(state="disabled")
        try:
            in_path = path_var.get().strip().strip('"')
            if not in_path or not os.path.isfile(in_path):
                messagebox.showwarning(APP_NAME, "입력 파일 경로가 올바르지 않습니다.", parent=root)
                return
            stem = os.path.splitext(os.path.basename(in_path))[0]
            out_path = filedialog.asksaveasfilename(
                parent=root,
                title="PPT 저장 위치",
                defaultextension=".pptx",
                initialfile=stem + ".pptx",
                filetypes=[("PowerPoint", "*.pptx")],
            )
            if not out_path:
                return
            gen_btn.config(text="생성 중…")
            root.config(cursor="wait")
            root.update_idletasks()
            try:
                engine2.generate(in_path, _template_path(), out_path, background=bg_var.get())
            except Exception as exc:  # dialog + error.log, never crash
                log = _log_error(exc)
                messagebox.showerror(APP_NAME, f"생성에 실패했습니다.\n\n{exc}\n\n로그: {log}", parent=root)
                return
            _open_folder(out_path)                       # 완료 후 저장 폴더 열기
            messagebox.showinfo(APP_NAME, f"완료: {out_path}", parent=root)
        finally:
            busy["on"] = False
            root.config(cursor="")
            gen_btn.config(state="normal", text="PPT 생성")

    def save_input_template():
        if busy["on"]:
            return
        busy["on"] = True
        form_btn.config(state="disabled")
        try:
            out = filedialog.asksaveasfilename(
                parent=root,
                title="입력 양식 저장",
                defaultextension=".xlsm",
                initialfile="입력양식.xlsm",
                filetypes=[("Excel 매크로 사용 통합 문서", "*.xlsm")],
            )
            if not out:
                return
            try:
                src = _input_template_path()
                with zipfile.ZipFile(src) as zf:        # embedded form must be intact (build check)
                    if zf.testzip() is not None:
                        raise ValueError("내장 양식이 손상되어 있습니다. 빌드를 다시 확인하세요.")
                shutil.copyfile(src, out)
                with zipfile.ZipFile(out) as zf:        # written file must be a valid zip
                    if zf.testzip() is not None:
                        raise ValueError("저장된 파일이 손상되었습니다. 동기화/백신 폴더가 아닌 "
                                         "로컬 폴더(예: C:\\Temp)에 저장해 보세요.")
            except Exception as exc:
                log = _log_error(exc)
                messagebox.showerror(APP_NAME, f"양식 저장에 실패했습니다.\n\n{exc}\n\n로그: {log}", parent=root)
                return
            _open_folder(out)
            messagebox.showinfo(APP_NAME, f"입력 양식 저장 완료:\n{out}", parent=root)
        finally:
            busy["on"] = False
            form_btn.config(state="normal")

    gen_btn = tk.Button(root, text="PPT 생성", width=16, height=1, command=generate)
    gen_btn.place(x=18, y=148)
    form_btn = tk.Button(root, text="입력 양식 저장", width=16, height=1, command=save_input_template)
    form_btn.place(x=170, y=148)

    tk.Label(root, text="처음이면 [입력 양식 저장]으로 양식(.xlsm)을 받아 작성한 뒤 변환하세요.",
             fg="#888888", font=("Segoe UI", 8)).place(x=18, y=218)

    root.mainloop()


def run_cli(argv):
    parser = argparse.ArgumentParser(description=f"{APP_NAME} {VERSION} (CLI mode)")
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--template", default=None)
    parser.add_argument("--background", default="none", choices=["none"])
    args = parser.parse_args(argv)
    try:
        summary = engine2.generate(args.input, args.template or _template_path(), args.output,
                                   background=args.background)
    except Exception as exc:  # noqa: BLE001
        log = _log_error(exc)
        print(f"ERROR: {exc} (log: {log})")
        sys.exit(1)
    print("OK:", summary["output"])
    for w in summary.get("warnings", []):
        print("WARN:", w)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        run_cli(sys.argv[1:])
    else:
        run_gui()
