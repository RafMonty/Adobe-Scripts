import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk
import filetype
import os

current_image = None
current_path = None

def load_image():
    """Load and inspect image"""
    global current_image, current_path
    file_path = filedialog.askopenfilename(title="Select an image")
    if not file_path:
        return

    try:
        kind = filetype.guess(file_path)
        file_info = []
        if kind:
            file_info.append(f"File Type: {kind.extension}")
            file_info.append(f"MIME Type: {kind.mime}")
        else:
            file_info.append("File Type: Unknown")

        img = Image.open(file_path)
        current_image = img
        current_path = file_path

        # Thumbnail
        thumb = img.copy()
        thumb.thumbnail((200, 200))
        photo = ImageTk.PhotoImage(thumb)
        thumbnail_label.config(image=photo)
        thumbnail_label.image = photo

        # Image info
        file_info.append(f"Format: {img.format}")
        file_info.append(f"Mode: {img.mode}")
        file_info.append(f"Size: {img.size[0]} × {img.size[1]}")
        file_info.append(f"Compression: {img.info.get('compression', 'N/A')}")
        file_info.append(f"Bits per pixel: {getattr(img, 'bits', 'N/A')}")
        file_info.append(f"ICC Profile: {'Yes' if 'icc_profile' in img.info else 'No'}")

        for k, v in img.info.items():
            file_info.append(f"{k}: {v}")

        file_size = os.path.getsize(file_path)
        file_info.append(f"File Size: {file_size / 1024:.2f} KB")

        text_box.config(state="normal")
        text_box.delete(1.0, tk.END)
        text_box.insert(tk.END, "\n".join(file_info))
        text_box.config(state="disabled")

        btn_save_fixed.config(state="normal")

    except Exception as e:
        messagebox.showerror("Error", f"Unable to load image.\n{e}")
        current_image = None
        current_path = None
        btn_save_fixed.config(state="disabled")


def save_fixed():
    """Save image as FIXED_ prefixed JPEG in same folder"""
    global current_image, current_path
    if not current_image or not current_path:
        messagebox.showwarning("No image", "Please load an image first.")
        return

    try:
        img_to_save = current_image.convert("RGB")
        folder = os.path.dirname(current_path)
        filename = os.path.basename(current_path)
        name, _ = os.path.splitext(filename)
        new_path = os.path.join(folder, f"FIXED_{name}.jpg")

        icc_profile = current_image.info.get("icc_profile")

        save_args = {"format": "JPEG", "quality": 70, "optimize": True}
        if icc_profile:
            save_args["icc_profile"] = icc_profile

        img_to_save.save(new_path, **save_args)
        messagebox.showinfo("Saved", f"Image saved:\n{new_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to save image.\n{e}")


# --- GUI layout ---
root = tk.Tk()
root.title("Image Info Inspector")
root.geometry("600x550")

frame = ttk.Frame(root, padding=10)
frame.pack(fill="both", expand=True)

# Buttons
btn_frame = ttk.Frame(frame)
btn_frame.pack(pady=5)

btn_load = ttk.Button(btn_frame, text="📂 Load Image", command=load_image)
btn_load.pack(side="left", padx=5)

btn_save_fixed = ttk.Button(btn_frame, text="💾 Save Fixed", command=save_fixed, state="disabled")
btn_save_fixed.pack(side="left", padx=5)

# Thumbnail + info
thumbnail_label = ttk.Label(frame)
thumbnail_label.pack(pady=5)

scroll_frame = ttk.Frame(frame)
scroll_frame.pack(fill="both", expand=True)

text_scrollbar = ttk.Scrollbar(scroll_frame)
text_scrollbar.pack(side="right", fill="y")

text_box = tk.Text(scroll_frame, wrap="word", yscrollcommand=text_scrollbar.set)
text_box.pack(fill="both", expand=True)
text_scrollbar.config(command=text_box.yview)
text_box.config(state="disabled")

root.mainloop()
