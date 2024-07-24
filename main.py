import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk

# Global variable untuk tree
tree = None

def tabel_nama(files):
    global tree

    def atur_tinggi_kolom(root, treeview):
        num_rows = len(treeview.get_children())
        total_height = num_rows * 25  # Misalkan setiap baris memiliki tinggi 25 piksel
        treeview.configure(height=num_rows)
        
        # Mendapatkan informasi layar
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        
        # Mengatur posisi jendela di tengah layar
        x_position = (screen_width - 500) // 2  # 500 adalah lebar jendela yang diinginkan
        y_position = (screen_height - total_height) // 2
        
        root.geometry(f"500x{total_height}+{x_position}+{y_position}")

    # Root window untuk tabel_nama
    root = tk.Tk()
    root.title("Tabel Nama File")

    frame = tk.Frame(root)
    frame.pack(pady=20, fill=tk.BOTH, expand=True)

    style = ttk.Style(root)
    style.configure("Treeview.heading", anchor="center")
    style.configure("Treeview", rowheight=25)

    # Membuat treeview
    tree = ttk.Treeview(frame, columns=("Customer", "Ukuran", "Qty"), show="headings")
    tree.heading("Customer", text="Customer", anchor="center")
    tree.heading("Ukuran", text="Ukuran", anchor="center")
    tree.heading("Qty", text="Qty", anchor="center")
    tree.pack(fill=tk.BOTH, expand=True)

    tree.column("Customer", width=150, anchor="center")
    tree.column("Ukuran", width=100, anchor="center")
    tree.column("Qty", width=80, anchor="center")

    for nama_files in files:
        split_str = nama_files.replace(".tif", "tif")
        pattern = re.compile(r"^(?:\d+\s*)?([\w\s]+?)\s*\(([\d\sXx]+)\)\s*(\d+)\s*(pcs|PCS)?\s*(?:[\w\s\+\-]+)?", re.IGNORECASE)
        match = pattern.match(split_str)

        if match:
            kolom1 = match.group(1)
            kolom2 = match.group(2)
            kolom3 = match.group(3)
            tree.insert("", tk.END, values=(kolom1, kolom2, kolom3))

    def on_copy(column_name):
        selected_items = tree.selection()
        data = []

        for item in selected_items:
            item_values = tree.item(item)["values"]
            data.append(str(item_values[tree["columns"].index(column_name)]))
            
        print(data)
        clip_text = "\n".join(data)
        root.clipboard_clear()
        root.clipboard_append(clip_text)
        messagebox.showinfo("Informasi", f"{column_name} Berhasil Disalin!")



    frameBtnClip = tk.Frame(root)
    frameBtnClip.pack(pady=10)

    btnClipCustomer = tk.Button(frameBtnClip, text="Customer", command=lambda: on_copy("Customer"))
    btnClipCustomer.pack(pady=10, side="left")
    btnClipUkuran = tk.Button(frameBtnClip, text="Ukuran", command=lambda: on_copy("Ukuran"))
    btnClipUkuran.pack(pady=10, side="left")
    btnClipQty = tk.Button(frameBtnClip, text="Qty", command=lambda: on_copy("Qty"))
    btnClipQty.pack(pady=10, side="left")

    atur_tinggi_kolom(root, tree)  # Memanggil fungsi untuk mengatur tinggi kolom

    root.mainloop()

def center_window(root, width, height):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")

def select_directory():
    directory_path = filedialog.askdirectory(title='Pilih Foldernya Lily...')
    if directory_path:
        files = list_file_onDirectory(directory_path)
        tabel_nama(files)
    else:
        print("Tidak ada direktori yang dipilih!")

def list_file_onDirectory(directory_path):
    if os.path.isdir(directory_path):
        return [f for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]
    else:
        print("Path tidak ada")
        return []

def mainProgram():
    root = tk.Tk()
    root.title("AMBIL NAMA FOLDER PAMULANG")
    width = 220
    height = 350
    center_window(root, width, height)

    # Menampilkan gambar
    image_path = os.path.join(os.path.dirname(__file__), "sasuke.png")
    image = Image.open(image_path)
    photo = ImageTk.PhotoImage(image)

    canvas = tk.Canvas(root, width=image.width, height=image.height, bg="white")
    canvas.pack()
    canvas.create_image(0, 0, anchor="nw", image=photo)
    canvas.image = photo

    lblTitle = tk.Label(root, text="MALAS NGETIK DI EXCEL!!!")
    lblTitle.pack()

    btnCariFolder = tk.Button(root, text=" Mulai!", command=select_directory)
    btnCariFolder.pack()

    root.mainloop()

if __name__ == "__main__":
    mainProgram()
