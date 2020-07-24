import xlwt

def up_b_factor_4():


    x1 = [-1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0,
     1.0, -2.0, 2.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    x2 = [-1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0,
     1.0, 0.0, 0.0, -2.0, 2.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    x3 = [-1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0,
     1.0, 0.0, 0.0, 0.0, 0.0, -2.0, 2.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    x4 = [-1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 
    1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -2.0, 2.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

    level_x1 = []
    level_x2 = []
    level_x3 = []
    level_x4 = []

    print('请分别依次输入四个因子最高水平')
    x1_h, x2_h, x3_h, x4_h = map(float, input().split())
    print('''请分别依次输入四个因子最低水平
    和因子高水平顺序一致''')
    x1_l, x2_l, x3_l, x4_l = map(float, input().split())
    print('四个因子水平分别为：',x1_h ,x1_l ,x2_h ,x2_l, x3_h, x3_l, x4_h, x4_l)

    h = 2
    l = -2

    def get_real_level(x,x_h,x_l):
        level_x = []
        for i in x:
            if i == 0:
                v = (x_h+x_l)/2
                level_x.append(v)
            elif i == 1:
                v = (((x_h-x_l)/2)/h)+(x_h+x_l)/2
                level_x.append(v)
            elif i == -1 :
                v = (x_h+x_l)/2-(((x_h-x_l)/2)/h)
                level_x.append(v)
            elif i == h:
                level_x.append(x_h)
            else :
                level_x.append(x_l)
        return level_x

    level_x1 = get_real_level(x1,x1_h,x1_l)
    level_x2 = get_real_level(x2,x2_h,x2_l)
    level_x3 = get_real_level(x3,x3_h,x3_l)
    level_x4 = get_real_level(x4,x4_h,x4_l)


    new_excel = xlwt.Workbook()
    new_sheet = new_excel.add_sheet("sheet1")

    style = xlwt.XFStyle()

    font = xlwt.Font()
    font.name = '宋体'
    font.boid = False
    font.height = 12*20
    style.font = font

    borders = xlwt.Borders()
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    style.borders = borders

    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_RIGHT
    alignment.vert = xlwt.Alignment.VERT_CENTER
    style.alignment = alignment

    new_sheet.write(0,0,'x1',style)
    new_sheet.write(0,1,'x2',style)
    new_sheet.write(0,2,'x3',style)
    new_sheet.write(0,3,'x4',style)
    new_sheet.write(0,4,'x1_real',style)
    new_sheet.write(0,5,'x2_real',style)
    new_sheet.write(0,6,'x3_real',style)
    new_sheet.write(0,7,'x4_real',style)
    new_sheet.write(0,8,'y1',style)
    new_sheet.write(0,9,'y2',style)

    for j in range(1,(len(x1)+1)):
        new_sheet.write(j,0,x1[j-1],style)
        new_sheet.write(j,1,x2[j-1],style)
        new_sheet.write(j,2,x3[j-1],style)
        new_sheet.write(j,3,x4[j-1],style)
        new_sheet.write(j,4,level_x1[j-1],style)
        new_sheet.write(j,5,level_x2[j-1],style)
        new_sheet.write(j,6,level_x3[j-1],style)
        new_sheet.write(j,7,level_x4[j-1],style)


    new_excel.save('D:/up_4factor.xlsx')
