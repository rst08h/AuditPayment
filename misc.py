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
            