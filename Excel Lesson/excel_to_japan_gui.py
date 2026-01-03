import os
import json
import threading
import pandas as pd
from gtts import gTTS
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


def build_lesson(input_excel, out_folder, log):
    try:
        df = pd.read_excel(input_excel).fillna("")
    except Exception as e:
        log(f"Error reading Excel: {e}")
        return

    lesson_name = os.path.splitext(os.path.basename(input_excel))[0]
    base = out_folder
    os.makedirs(base, exist_ok=True)
    audio_dir_name = f"audio_{lesson_name}"
    audio_dir = os.path.join(base, audio_dir_name)
    os.makedirs(audio_dir, exist_ok=True)

    lesson_data = []
    mapping = []

    for i, r in df.iterrows():
        row_list = r.tolist()
        while len(row_list) < 4:
            row_list.append("")
        row = [str(x) for x in row_list[:4]]
        lesson_data.append(row)

        text = row[0].strip()
        if text:
            g_file = f"gtts_a_{i}.mp3"
            g_path = os.path.join(audio_dir, g_file)
            try:
                tts = gTTS(text, lang='ja')
                tts.save(g_path)
                mapping.append({
                    "id": f"a_{i}",
                    "text": text,
                    "file": f"{audio_dir_name}/{g_file}",
                    "lang": "ja",
                    "type": "a",
                    "engine": "gTTS",
                    "voice": "gTTS-ja"
                })
                log(f"Saved audio: {g_path}")
            except Exception as e:
                log(f"Warning: failed to generate audio for row {i}: {e}")

    lesson_file = os.path.join(base, f"{lesson_name}.json")
    try:
        with open(lesson_file, "w", encoding="utf-8") as f:
            json.dump(lesson_data, f, ensure_ascii=False, indent=2)
        log(f"Wrote lesson: {lesson_file}")
    except Exception as e:
        log(f"Error writing lesson json: {e}")

    mapping_file = os.path.join(base, f"mapping_{lesson_name}.json")
    try:
        with open(mapping_file, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
        log(f"Wrote mapping: {mapping_file}")
    except Exception as e:
        log(f"Error writing mapping json: {e}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel -> Japanese LESSON (GUI)")
        self.geometry("720x480")

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        # Input file / directory
        row = ttk.Frame(frm)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text="Input file or directory:").pack(side=tk.LEFT)
        self.input_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.input_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(row, text="Browse File", command=self.browse_file).pack(side=tk.LEFT)
        ttk.Button(row, text="Browse Dir", command=self.browse_dir).pack(side=tk.LEFT, padx=4)

        # Process all checkbox
        row2 = ttk.Frame(frm)
        row2.pack(fill=tk.X)
        self.all_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(row2, text="Process all .xlsx in directory", variable=self.all_var).pack(side=tk.LEFT)

        # Output folder
        row3 = ttk.Frame(frm)
        row3.pack(fill=tk.X, pady=4)
        ttk.Label(row3, text="Output folder:").pack(side=tk.LEFT)
        self.out_var = tk.StringVar(value=os.path.join(os.path.dirname(__file__), "..", "LESSON"))
        ttk.Entry(row3, textvariable=self.out_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(row3, text="Browse", command=self.browse_out).pack(side=tk.LEFT)

        # Controls
        ctl = ttk.Frame(frm)
        ctl.pack(fill=tk.X, pady=6)
        self.start_btn = ttk.Button(ctl, text="Start", command=self.start)
        self.start_btn.pack(side=tk.LEFT)
        self.stop_requested = False

        # Log
        ttk.Label(frm, text="Log:").pack(anchor=tk.W)
        self.logbox = scrolledtext.ScrolledText(frm, height=15)
        self.logbox.pack(fill=tk.BOTH, expand=True)

    def browse_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if p:
            self.input_var.set(p)

    def browse_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.input_var.set(p)

    def browse_out(self):
        p = filedialog.askdirectory()
        if p:
            self.out_var.set(p)

    def log(self, msg):
        self.logbox.insert(tk.END, msg + "\n")
        self.logbox.see(tk.END)

    def start(self):
        inp = self.input_var.get().strip()
        out = self.out_var.get().strip()
        if not inp:
            messagebox.showwarning("Input missing", "Please choose an input file or directory.")
            return
        if not out:
            messagebox.showwarning("Output missing", "Please choose an output folder.")
            return

        os.makedirs(out, exist_ok=True)
        self.start_btn.config(state=tk.DISABLED)
        self.stop_requested = False
        t = threading.Thread(target=self._run, args=(inp, out, self.all_var.get()), daemon=True)
        t.start()

    def _run(self, inp, out, all_mode):
        try:
            if all_mode and os.path.isdir(inp):
                files = [os.path.join(inp, f) for f in os.listdir(inp) if f.lower().endswith('.xlsx')]
                if not files:
                    self.log(f"No .xlsx files in {inp}")
                for f in files:
                    if self.stop_requested:
                        break
                    self.log(f"Processing: {f}")
                    build_lesson(os.path.abspath(f), out, self.log)
            else:
                if not os.path.exists(inp):
                    self.log(f"Input not found: {inp}")
                else:
                    build_lesson(os.path.abspath(inp), out, self.log)
            self.log("All done.")
            messagebox.showinfo("Done", "Processing finished.")
        finally:
            self.start_btn.config(state=tk.NORMAL)


if __name__ == '__main__':
    app = App()
    app.mainloop()
