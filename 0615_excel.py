import MySQLdb
import pdb
import xlwt
import datetime
import time


import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


def export_excel(results,results1):
    # 导出数据
    # 创建workbook 设置编码
    workbook = xlwt.Workbook(encoding = 'utf-8')
    # 创建sheet
    re_sheet = workbook.add_sheet('receiving')
    # bill_sheet
    bill_sheet = workbook.add_sheet('bill')
    # 写表头
    receiving_keys = ['发货方','收货方','批次号','箱号','OMS PO号',
    'SKU','预计收货件数','实际收货件数','货物到仓日期','收货完成日期','实际收货时效(天)','仓库备注']
    # bill_keys
    bill_keys = ['快递送达时间','快递单号','退货订单编号','收货店铺','SKU','订单数量(件)','质检情况']
    # 

    # 设置表头对齐方式
    style_head = xlwt.XFStyle()  # 创建一个样式对象，初始化样式
    al_head = xlwt.Alignment()
    al_head.horz = 0x02      # 设置水平居中
    al_head.vert = 0x01      # 设置垂直居中
    style_head.alignment = al_head
    

    # 设置背景颜色
    pt_head = xlwt.Pattern()
    # 设置背景颜色的模式
    pt_head.pattern = xlwt.Pattern.SOLID_PATTERN
    # 背景颜色
    pt_head.pattern_fore_colour = 24
    style_head.pattern = pt_head

    # 表头字体设置
    pt_ft = xlwt.Font()
    pt_ft.bold = True   #加粗
    pt_ft.height = 30*11
    style_head.font = pt_ft
    
    tall_style = xlwt.easyxf('font:height 720')  # 表头高设计（表头的下标是0）
    re_sheet.row(0).set_style(tall_style)

    # 设置收货表单正文样式
    style_re = xlwt.XFStyle()  # 创建一个样式对象，初始化样式

    # 设置正文字体
    ft_re = xlwt.Font()
    ft_re.height=20*11
    style_re.font=ft_re
    
    # 设置列宽
    for re in receiving_keys:
        re_sheet.col(receiving_keys.index(re)).width = 11 * 500   # 循环设置每列的宽都是一样的

    re_sheet.write_merge(0,0,0,len(receiving_keys)-1,'收货报表',style_head)
    bill_sheet.write_merge(0,0,0,len(bill_keys)-1,'订单信息')

    
    path = 'C:\\Users\\Administrator\\\Desktop\\0615\\rece'
    time = int(datetime.datetime.now().timestamp())
    path = path+str(time)+'.xls'
    
    
    for key in receiving_keys:
        re_sheet.write(1,receiving_keys.index(key),str(key),style_re)
    
    for key in bill_keys:
        bill_sheet.write(1,bill_keys.index(key),str(key))

    row = 2
    # 保存正文receiving表的正文
    for re in results:
        col=0
        for r in re:
            # id 不用导出，所以跳过
            if re.index(r) == 0:
                continue
            re_sheet.write(row,col,r,style_re)    # 这里加上了re表正文的格式
            col = col+1
        row = row+1


    row = 2
    # 保存bill表的正文
    for re in results1:
        col=0
        for r in re:
            # id 不用导出，所以跳过
            if re.index(r) == 0:
                continue
            bill_sheet.write(row,col,r)
            col = col+1
        row = row+1
    print('保存成功')
   
 

    workbook.save(path)

    send_email(path)

def send_email(path):
    fromaddr = '846848165@qq.com'
    password = 'kejkxljzxvurbdfj'
    toaddrs = ['1248773869@qq.com']


    content = 'hello, this is email content.'
    textApart = MIMEText(content)

    excelApart = MIMEApplication(open(path, 'rb').read())
    excelApart.add_header('Content-Disposition', 'attachment', filename=path)


    m = MIMEMultipart()
    m.attach(textApart)
    m.attach(excelApart)
   

    try:
        server = smtplib.SMTP('smtp.qq.com')
        server.login(fromaddr,password)
        server.sendmail(fromaddr, toaddrs, m.as_string())
        print('success')
        server.quit()
    except smtplib.SMTPException as e:
        print('error:',e) #打印错误


conn= MySQLdb.connect(
    host='localhost',
    port = 3306,
    user='root',
    passwd='123456',
    db ='goods', #数据库名
    charset='utf8' # 避免中文乱码
    )

cur = conn.cursor()
cur.execute("select * from receiving") # 执行查询
# ((1, '天猫', '个人', 'Mrqq', 'werwe', '234345', '456gdf', 45, 45, '2020-06-01', '2020-06-05', '5', '无'),)
result = cur.fetchall()
cur.execute('select * from bill')
result1 = cur.fetchall()
export_excel(result,result1)
