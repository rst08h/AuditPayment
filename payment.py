import screen_config
import tkinter as tk
from tkinter import filedialog as fd
import pandas as pd
import threading
import time
from datetime import datetime


def dateparse(dates): return [
    datetime.strptime(d, '%d.%m.%Y') for d in dates]


def process_file1():
    if len(txtbox1.get("1.0", tk.END)) >= 2:
        txtbox1.delete("1.0", tk.END)
    txtbox1.insert(tk.END, '[ประมวลผลไฟล์]\n')
    if filename1.cget('text') != "[ไม่ได้เลือก]":
        saptextfile = filename1.cget('text')
        saptextfile2 = saptextfile+'.txt'
        file1 = open(saptextfile, 'r', encoding='utf_16_le')
        file2 = open(saptextfile2, 'w', encoding='utf-8')
        file1.read(1)
        line = file1.readline()
        if line.find('ZFAPRP08') > 0:
            txtbox1.insert(tk.END, 'อ่านไฟล์ ' + saptextfile + '\n')
            while line != '':
                # print(line)
                if (line[0] != '0') and (line[0] != '1') and (line[0] != '2') and (line[0] != '3') and (line[0] != '#'):
                    file2.write('#'+line)
                else:
                    file2.write(line)
                line = file1.readline()
            file1.close()
            file2.close()
            def dateparse(x): return datetime.strptime(x, '%Y-%m-%d %H:%M:%S')

            cols = [0, 1, 6, 11, 12, 16, 19, 21]
            colsname = ['posting_date', 'doc_no', 'tax_no', 'doctype',
                        'payee_name', 'description', 'amount', 'collected']

            df = pd.read_table(saptextfile2, sep='\t', engine='python', date_parser=dateparse, parse_dates=[
                               'posting_date'], header=None, usecols=cols, names=colsname, comment='#', dtype={'doctype': 'category', 'collected': 'category'}, encoding='utf-8')
            df.doc_no = df.doc_no.astype('str')
            df['amount'] = df['amount'].str.replace(',', '')
            df['amount'] = df['amount'].str.replace('.', '')
            df.amount = df.amount.astype('int')  # ตอนที่แสดงผลต้องหารด้วย 100
            df.to_excel(saptextfile+'.xlsx', index=False)
            txtbox1.insert(tk.END, 'ส่งออก ZFAPRP08 ไปยัง ' +
                           saptextfile+'.xlsx\n')

            # วิเคราะห์ข้อมูล
            txtbox1.insert(tk.END, '[วิเคราะห์ข้อมูล]\n')
            txtbox1.insert(tk.END, 'วันที่ ')
            
        else:
            print("ไม่ใช่ไฟล์ ZFAPRP08")
            txtbox1.insert(tk.END, 'ไม่ใช่ไฟล์ ZFAPRP08\n')
    submit_btn['state'] = tk.NORMAL
    txtbox1.insert(tk.END, '[จบกระบวนการ]\n')


def btn1_command():
    f = fd.askopenfilename()
    filename1.config(text=f)


def btn2_command():
    f = fd.askopenfilename()
    filename2.config(text=f)


def process_command():
    submit_btn['state'] = tk.DISABLED
    athread = threading.Thread(target=process_file1)
    athread.start()
    # process_file1()
    # proc_window=tk.Toplevel(window)
    # proc_window.title="ประมวลผล"


window = tk.Tk()
window.title("เครื่องมือตรวจสอบ การจ่ายเงินสดย่อย โดยสำนักตรวจสอบกระบวนการหลัก")
print(window.geometry())
# print(screen_config.Get_HWND_DPI(window.winfo_id()))
# screen_config.MakeTkDPIAware(window)
s = screen_config.scaling(window.winfo_id())
#print('scaling : ' + s)
window.tk.call('tk', 'scaling', s)


print(window.geometry())
upper_panel1 = tk.Frame(window, padx=2, pady=2)
upper_panel1.pack()
upper_panel2 = tk.Frame(window, padx=10, pady=2)
upper_panel2.pack()
upper_panel3 = tk.Frame(window, padx=2, pady=2)
upper_panel3.pack()
lower_panel = tk.Frame(window, padx=5, pady=5)
lower_panel.pack()

tk.Label(upper_panel1, text="ไฟล์ ZFAPRP08 :").pack(side="left")
filename1 = tk.Label(upper_panel1, text="[ไม่ได้เลือก]")
filename1.pack(side='left')
tk.Button(upper_panel1, text="เลือกไฟล์",
          command=btn1_command).pack(side='left')
tk.Label(upper_panel2, text="ไฟล์ GL :").pack(side='left')
filename2 = tk.Label(upper_panel2, text="[ไม่ได้เลือก]")
filename2.pack(side='left')
tk.Button(upper_panel2, text="เลือกไฟล์",
          command=btn2_command).pack(side="left")
submit_btn = tk.Button(upper_panel3, text="ดำเนินการ", command=process_command)
submit_btn.pack()

txtbox1 = tk.Text(lower_panel)
txtbox1.grid(row=0, column=0)
txtbox2 = tk.Text(lower_panel)
txtbox2.grid(row=0, column=1)

# print(fd.askopenfile())

window.mainloop()
