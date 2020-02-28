import tkinter as tk

window = tk.Tk()
window.title('TRPG replay maker')

label_fp = tk.Label(window, 
    text='LOG文件路径',         # 标签的文字
    font=('Arial', 12),         # 字体和字体大小
    )
label_fp.grid(row=0, column=0)

entry_fp = tk.Entry(window)
entry_fp.grid(row=0, column=1)


label_fp = tk.Button(window, 
    text='保存',         # 标签的文字
    font=('Arial', 12),         # 字体和字体大小
    )
label_fp.grid(row=0, column=2)

entry_fp = tk.Entry(window)
entry_fp.grid(row=1, column=1)

window.mainloop()