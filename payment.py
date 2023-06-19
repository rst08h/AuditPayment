import tkinter as tk
window = tk.Tk()
window.title("กระดาษทำการ การจ่ายเงิน")


tk.Label(text="ไฟล์ที่ 1").grid(row=0, column=0)
tk.Entry().grid(row=0, column=1)
tk.Button(text="เลือกไฟล์").grid(row=0, column=2)
tk.Label(text="ไฟล์ที่ 2").grid(row=1, column=0)
tk.Entry().grid(row=1, column=1)
tk.Button(text="เลือกไฟล์").grid(row=1, column=2)
tk.Button(text="ดำเนินการ").grid(row=2, column=1)

window.mainloop()
