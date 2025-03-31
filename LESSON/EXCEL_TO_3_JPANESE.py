import os
import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from edge_tts import Communicate
import asyncio

# Hàm chọn file Excel
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

# Hàm lấy tên bài học từ tên file
def get_lesson_name(file_path):
    return os.path.splitext(os.path.basename(file_path))[0]

# Hàm chuyển đổi từ full-width sang half-width
def normalize_numbers(text):
    return "".join(chr(ord(char) - 0xFEE0) if 0xFF10 <= ord(char) <= 0xFF19 else char for char in text)

# Hàm tạo âm thanh và file mapping JSON với Edge_TTS
async def create_audio_and_mapping(df, lesson_name, voice_code, suffix, lang):
    output_dir = f"audio_{lesson_name}_{suffix}" if suffix else f"audio_{lesson_name}"
    os.makedirs(output_dir, exist_ok=True)

    words_col0 = df.iloc[:, 0].dropna().astype(str).tolist()
    words_col3 = df.iloc[:, 3].dropna().astype(str).tolist()
    
    audio_mapping = []
    for idx, (word0, word3) in enumerate(zip(words_col0, words_col3)):
        # Chuẩn hóa dấu nháy và chuẩn hóa số
        word0 = normalize_numbers(word0.replace("’", "'").replace("“", '"').replace("”", '"'))
        word3 = normalize_numbers(word3.replace("’", "'").replace("“", '"').replace("”", '"'))

        # Âm thanh cột 0
        filename0 = f"{output_dir}/audio_{lesson_name}_{idx + 1}_col0.mp3"
        await Communicate(word0, voice_code).save(filename0)
        audio_mapping.append({"text": word0, "file": filename0})

        # Âm thanh cột 3
        filename3 = f"{output_dir}/audio_{lesson_name}_{idx + 1}_col3.mp3"
        await Communicate(word3, voice_code).save(filename3)
        audio_mapping.append({"text": word3, "file": filename3})

    mapping_file = f"mapping_{lesson_name}_{suffix}.json" if suffix else f"mapping_{lesson_name}.json"
    with open(mapping_file, "w", encoding="utf-8") as f:
        json.dump(audio_mapping, f, ensure_ascii=False, indent=4)

    return output_dir, mapping_file

# Hàm tạo file JSON 6 cột
def create_lesson_json(df, lesson_name, suffix):
    df = df.applymap(lambda x: normalize_numbers(str(x).replace("’", "'")) if isinstance(x, str) else str(x))
    data = df.iloc[:, :6].fillna("").values.tolist()
    lesson_json = f"{lesson_name}_{suffix}.json" if suffix else f"{lesson_name}.json"
    with open(lesson_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    return lesson_json

# Hàm xử lý chính khi bấm nút "Tạo Tất Cả"
async def generate_all():
    file_path = entry_file.get()
    if not file_path:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel!")
        return
    
    voice_choice = voice_var.get()
    lang_choice = lang_var.get()

    voice_mapping = {
        "Male": ("en-US-GuyNeural", "m"),
        "Female": ("en-US-JennyNeural", "f"),
        "Child": ("en-US-AriaNeural", "c"),
        "Japanese": ("ja-JP-NanamiNeural", "")  # Bỏ "j" khỏi suffix
    }

    lang_mapping = {
        "English": "en",
        "Japanese": "ja"
    }

    if voice_choice not in voice_mapping or lang_choice not in lang_mapping:
        messagebox.showerror("Lỗi", "Vui lòng chọn giọng đọc và ngôn ngữ!")
        return

    voice_code, suffix = voice_mapping[voice_choice]
    lang = lang_mapping[lang_choice]
    
    try:
        df = pd.read_excel(file_path)
        lesson_name = get_lesson_name(file_path)

        if df.shape[1] < 6:
            messagebox.showerror("Lỗi", "File Excel phải có ít nhất 6 cột!")
            return

        # 1. Tạo âm thanh và mapping
        output_dir, mapping_file = await create_audio_and_mapping(df, lesson_name, voice_code, suffix, lang)
        
        # 2. Tạo file JSON 6 cột
        lesson_json = create_lesson_json(df, lesson_name, suffix)

        messagebox.showinfo(
            "Hoàn tất",
            f"Đã tạo thành công:\n- Thư mục âm thanh: {output_dir}\n- Mapping JSON: {mapping_file}\n- Lesson JSON: {lesson_json}"
        )
    except Exception as e:
        messagebox.showerror("Lỗi", f"Lỗi khi thực hiện: {str(e)}")

# Giao diện GUI
root = tk.Tk()
root.title("Chương Trình 3 Trong 1: Audio - Mapping - Lesson JSON (Edge_TTS)")
root.geometry("600x350")

# Label + Entry để chọn file Excel
tk.Label(root, text="File Excel:").pack(pady=5)
frame_file = tk.Frame(root)
frame_file.pack()
entry_file = tk.Entry(frame_file, width=60)
entry_file.pack(side=tk.LEFT, padx=5)
btn_select = tk.Button(frame_file, text="Chọn File", command=select_file)
btn_select.pack(side=tk.LEFT)

# Lựa chọn giọng đọc
tk.Label(root, text="Chọn Giọng Đọc:").pack(pady=5)
voice_var = tk.StringVar(value="Male")
voice_choices = ["Male", "Female", "Child", "Japanese"]
voice_menu = ttk.Combobox(root, textvariable=voice_var, values=voice_choices)
voice_menu.pack(pady=5)

# Lựa chọn ngôn ngữ
tk.Label(root, text="Chọn Ngôn Ngữ:").pack(pady=5)
lang_var = tk.StringVar(value="English")
lang_choices = ["English", "Japanese"]
lang_menu = ttk.Combobox(root, textvariable=lang_var, values=lang_choices)
lang_menu.pack(pady=5)

# Nút thực hiện tất cả
btn_generate = tk.Button(root, text="Tạo Tất Cả", command=lambda: asyncio.run(generate_all()), bg="green", fg="white")
btn_generate.pack(pady=20)

# Chạy giao diện
root.mainloop()
