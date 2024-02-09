def thaidate(td): 
    day=int(td.strftime('%d'))
    month=td.strftime('%m')
    match month:
        case '01':
            m_name='ม.ค.'
        case '02':
            m_name='ก.พ.'
        case '03':
            m_name='มี.ค.'
        case '04':
            m_name='เม.ย.'
        case '05':
            m_name='พ.ค.'
        case '06':
            m_name='มิ.ย.'
        case '07':
            m_name='ก.ค.'
        case '08':
            m_name='ส.ค.'
        case '09':
            m_name='ก.ย.'
        case '10':
            m_name='ต.ค.'
        case '11':
            m_name='พ.ย.'
        case '12':
            m_name='ธ.ค.'
    
    B_year=int(td.strftime('%Y'))+543
    return str(day) + ' ' +  m_name + ' ' + str(B_year)
            
def thaiym(td): 
    month=td.strftime('%m')
    match month:
        case '01':
            m_name='ม.ค.'
        case '02':
            m_name='ก.พ.'
        case '03':
            m_name='มี.ค.'
        case '04':
            m_name='เม.ย.'
        case '05':
            m_name='พ.ค.'
        case '06':
            m_name='มิ.ย.'
        case '07':
            m_name='ก.ค.'
        case '08':
            m_name='ส.ค.'
        case '09':
            m_name='ก.ย.'
        case '10':
            m_name='ต.ค.'
        case '11':
            m_name='พ.ย.'
        case '12':
            m_name='ธ.ค.'
    
    B_year=int(td.strftime('%Y'))+543
    return m_name + ' ' + str(B_year)


def cut_tab(line):
    linetemp=""
    adata=[]
    seperate=0
    for ch in line:
        if ch != '\t' and ch != '\n':
            if seperate==1:
                seperate=0
            linetemp=linetemp + ch
            tt=0  
        else:
            if seperate==0:
                adata.append(linetemp)
                linetemp=''

            seperate=1
            tt=tt+1
            if len(adata)==2 and tt == 6 :
                adata.append('<ไม่ได้ระบุ>')
                tt=0
            if len(adata)==5 and tt==5:
                adata.append('<ไม่ได้ระบุ>')
                tt=0
    return adata


def line_clinsing(line):
    result=""
    adata=cut_tab(line)
    if len(adata) == 9 :
        adata[5]=adata[5]+adata[6] # รวมกรณี description มีแทปคั่นกลาง
        adata.remove(adata[6])
    
    #จัดรูปแบบฟิลจำนวนเงิน
    linetemp=""
    for ch in adata[6]:
        if ch == ' ' or ch == ',' or ch == '"' : #ฟิวจำนวนเงินมักมีช่องว่างข้างหน้า และจะไม่เอาคอมม่ากับฟันหนูด้วย
            print('not use')
        else:
            linetemp=linetemp + ch
    adata[6]=linetemp

    for linetmp in adata:
        result=result+linetmp+'\t'

    result=result+'\n'
    return result    
