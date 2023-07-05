import screen_config
import tkinter as tk
from tkinter import filedialog as fd
import pandas as pd
import threading
import time
from datetime import datetime
import misc
from decimal import Decimal
from decimal import getcontext

getcontext().prec=2
 # บน Mac ไม่สามารถใช้ฟังชั่นของ excel ใน pandas พร้อมกันได้ 
 # Flag นี้เอาไว้ทำ process syncro 
xl_flag = 0

###################################################################################
#
#   PROCESS FILE 1 ไฟล์ ZFAPRP08 โดย User export มาเป็น .xls แต่จริงๆ ข้าในเป็น TEXT FILE utf_16_le
#
###################################################################################
def process_file1():
    if len(txtbox1.get("1.0", tk.END)) >= 2:
        txtbox1.delete("1.0", tk.END)
    
    if filename1.cget('text') != "[ไม่ได้เลือก]":
        txtbox1.insert(tk.END, '[ประมวลผลไฟล์]\n')
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
            #def dateparse(x): return datetime.strptime(x, '%Y-%m-%d %H:%M:%S')

            def dateparse(dates): return [
                datetime.strptime(d, '%d.%m.%Y') for d in dates]
            cols = [0, 1, 6, 11, 12, 16, 19, 21]
            colsname = ['posting_date', 'doc_no', 'tax_no', 'doctype',
                        'payee_name', 'description', 'amount', 'collected']

            df1 = pd.read_table(saptextfile2, sep='\t', engine='python', date_parser=dateparse, parse_dates=[
                'posting_date'], header=None, usecols=cols, names=colsname, comment='#', dtype={'doctype': 'category', 'collected': 'category'}, encoding='utf-8')
            df1.doc_no = df1.doc_no.astype('str')
            df1['amount'] = df1['amount'].str.replace(',', '')
            #df1['amount'] = df1['amount'].str.replace('.', '')
            # ตอนที่แสดงผลต้องหารด้วย 100
            #df1.amount = df1.amount.astype('int')
            global xl_flag
            while xl_flag != 0: # proccess sync
                time.sleep(0.1)
            xl_flag = 1
            df1.to_excel(saptextfile+'.xlsx', index=False)
            txtbox1.insert(tk.END, 'ส่งออก ZFAPRP08 ไปยัง ' + saptextfile+'.xlsx\n')
            xl_flag = 0
                

            ### วิเคราะห์ข้อมูล
            txtbox1.insert(tk.END, '[วิเคราะห์ข้อมูล]\n')
            # จำนวนรายการ
            record_count=df1['posting_date'].count()
            txtbox1.insert(tk.END,'จำนวนรายการ : ' + str(record_count) + ' รายการ\n' )
            
            # ช่วงของวันที่ผ่านรายการ
            dmin = df1['posting_date'].min()
            dmax = df1['posting_date'].max()
            txtbox1.insert(tk.END, 'วันที่ผ่านรายการ ระหว่างวันที่ ' + misc.thaidate(dmin) +
                           ' ถึงวันที่ '+misc.thaidate(dmax) + '\n')
            # รวมฝั่งบวก และ ลบ
            positive_sum = Decimal('0')
            negative_sum = Decimal('0')
            for i in df1['amount']:
                if Decimal(i) > 0:
                    positive_sum+=Decimal(i)
                else:
                    negative_sum+=Decimal(i)
            txtbox1.insert(tk.END,'จำนวนเงินฝั่งบวก = {:,.2f} บาท\n'.format(positive_sum))
            txtbox1.insert(tk.END,'จำนวนเงินฝั่งลบ = {:,.2f} บาท\n'.format(negative_sum)) 
            txtbox1.insert(tk.END,'ผลต่าง {:,.2f} บาท\n'.format(positive_sum+negative_sum))
            



            

        else:
            print("ไม่ใช่ไฟล์ ZFAPRP08")
            txtbox1.insert(tk.END, 'ไม่ใช่ไฟล์ ZFAPRP08\n')
    submit_btn['state'] = tk.NORMAL
    txtbox1.insert(tk.END, '[จบกระบวนการ]\n')

###################################################################################
#
#   PROCESS FILE 2 ไฟล์ GL ที่เป็น Excel อยู่แล้ว
#
###################################################################################


def process_file2():
    if len(txtbox2.get("1.0", tk.END)) >= 2:
        txtbox2.delete("1.0", tk.END)
    
    if filename2.cget('text') != "[ไม่ได้เลือก]":
        txtbox2.insert(tk.END, '[ประมวลผลไฟล์]\n')
        glfilename = filename2.cget('text')
        cols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
        colsname = ['อ้างอิง', 'การกำหนด', 'เลขที่เอกสาร', 'วันที่เอกสาร', 'วันที่ผ่านรายการ', 'คีย์การผ่านรายการ',
                    'ประเภทเอกสาร', 'ชื่อผู้ใช้', 'เขตธุรกิจ', 'จำนวนเงิน', 'เอกสารกาหักล้าง', 'วันที่หักล้าง', 'ข้อความ']
        txtbox2.insert(tk.END, 'อ่านไฟล์ ' + glfilename + '\n')
        global xl_flag
        while  xl_flag != 0: # proccess sync
            time.sleep(0.1)
        xl_flag = 1
        df2 = pd.read_excel(glfilename, names=colsname,
                            usecols=cols, dtype={'เลขที่เอกสาร': 'str','จำนวนเงิน':'str'})
        xl_flag = 0
        ## กรองเอาเฉพาะ 'บันทึกชดเชยเงินสดย่อยจากรายได้'
        df2=df2[df2['ข้อความ']=='บันทึกชดเชยเงินสดย่อยจากรายได้']

        ### วิเคราะห์ข้อมูล
        txtbox2.insert(tk.END, '[วิเคราะห์ข้อมูล]\n')
        # จำนวนรายการ
        record_count=df2['วันที่ผ่านรายการ'].count()
        txtbox2.insert(tk.END,'จำนวนรายการที่เป็น "บันทึกชดเชยเงินสดย่อยจากรายได้" : ' + str(record_count) + ' รายการ\n' )
        
        # ช่วงของวันที่ผ่านรายการ
        dmin = df2['วันที่ผ่านรายการ'].min()
        dmax = df2['วันที่ผ่านรายการ'].max()
        txtbox2.insert(tk.END, 'วันที่ผ่านรายการ ระหว่างวันที่ ' + misc.thaidate(dmin) +
                        ' ถึงวันที่ '+misc.thaidate(dmax) + '\n')
      # รวมฝั่งบวก และ ลบ
        positive_sum = Decimal('0')
        negative_sum = Decimal('0')
        for i in df2['จำนวนเงิน']:
            if Decimal(i) > 0:
                positive_sum+=Decimal(i)
            else:
                negative_sum+=Decimal(i)
        txtbox2.insert(tk.END,'จำนวนเงินฝั่งบวก = {:,.2f} บาท\n'.format(positive_sum))
        txtbox2.insert(tk.END,'จำนวนเงินฝั่งลบ = {:,.2f} บาท\n'.format(negative_sum))
        txtbox2.insert(tk.END,'ผลต่าง {:,.2f} บาท\n'.format(positive_sum+negative_sum))
    
        txtbox2.insert(tk.END, '[จบกระบวนการ]\n')

def btn1_command():
    f = fd.askopenfilename()
    filename1.config(text=f)


def btn2_command():
    f = fd.askopenfilename()
    filename2.config(text=f)


def process_command():
    submit_btn['state'] = tk.DISABLED
    thread1 = threading.Thread(target=process_file1)
    thread2 = threading.Thread(target=process_file2)
    thread1.start()
    thread2.start()

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

tk.Label(lower_panel,text='ZFAPRP08').grid(row=0,column=0)
txtbox1 = tk.Text(lower_panel,font=('Tahoma',12))
txtbox1.grid(row=1, column=0)
tk.Label(lower_panel,text='GL').grid(row=0,column=1)
txtbox2 = tk.Text(lower_panel,font=('Tahoma',12))
txtbox2.grid(row=1, column=1)

# print(fd.askopenfile())

window.mainloop()
