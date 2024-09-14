import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import win32file
import win32api
import ctypes

def eject_drive(drive_letter):
    drive_path = f"\\\\.\\{drive_letter}:"
    
    try:
        handle = win32file.CreateFile(
            drive_path,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE,
            None,
            win32file.OPEN_EXISTING,
            0,
            None
        )
    except win32api.error as e:
        return False, f"无法访问驱动器 {drive_letter}: {e}"

    if handle == win32file.INVALID_HANDLE_VALUE:
        return False, f"无法获取驱动器 {drive_letter} 的句柄"

    try:
        ctypes.windll.kernel32.DeviceIoControl(
            handle.handle,
            0x2D4808,  # IOCTL_STORAGE_EJECT_MEDIA
            None,
            0,
            None,
            0,
            ctypes.byref(ctypes.c_ulong()),
            None
        )
        return True, f"{drive_letter} 盘已成功弹出。"
    except win32api.error as e:
        return False, f"弹出驱动器 {drive_letter} 时出错: {e}"
    finally:
        win32file.CloseHandle(handle)

def on_eject():
    drive_letter = drive_var.get()
    status_label.config(text="正在尝试弹出U盘...")
    root.update()
    
    success, message = eject_drive(drive_letter)
    status_label.config(text=message)
    
    if success:
        messagebox.showinfo("成功", message)
    else:
        messagebox.showerror("错误", message)

root = tk.Tk()
root.title("U盘弹出工具")
root.geometry("400x300")
root.configure(bg='#f0f0f0')

style = ttk.Style()
style.theme_use('clam')

style.configure('TFrame', background='#f0f0f0')
style.configure('TLabel', background='#f0f0f0', font=('Arial', 12))
style.configure('TButton', font=('Arial', 12, 'bold'))
style.map('TButton',
    background=[('active', '#4CAF50'), ('!active', '#45a049')],
    foreground=[('active', 'white'), ('!active', 'white')])

main_frame = ttk.Frame(root, padding="20 20 20 20")
main_frame.pack(expand=True, fill='both')

title_label = ttk.Label(main_frame, text="U盘弹出工具", font=('Arial', 18, 'bold'))
title_label.pack(pady=(0, 20))

drive_frame = ttk.Frame(main_frame)
drive_frame.pack(fill='x', pady=10)

drive_label = ttk.Label(drive_frame, text="选择要弹出的驱动器:")
drive_label.pack(side='left')

drive_var = tk.StringVar(value="E")
drive_menu = ttk.Combobox(drive_frame, textvariable=drive_var, values=["C", "D", "E", "F", "G", "H"], width=5, state="readonly")
drive_menu.pack(side='left', padx=(10, 0))

eject_button = ttk.Button(main_frame, text="弹出U盘", command=on_eject, style='TButton')
eject_button.pack(pady=20)

status_label = ttk.Label(main_frame, text="")
status_label.pack()

# 添加一个装饰性的分隔线
ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=20)

# 添加一个版权信息
copyright_label = ttk.Label(main_frame, text="© 2024 王鹏 Base on Claude-3.5-Sonnet", font=('Arial', 8))
copyright_label.pack(side='bottom')

root.mainloop()
