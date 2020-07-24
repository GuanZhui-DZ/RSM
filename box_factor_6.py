import xlwt

def box_b_factor_6():


    x1 = [-1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 
    0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 
    1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 
    1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    x2 = [-1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 1.0,
     1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
      0.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 
      0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]
    x3 = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 
    1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 
    0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, 0.0, 
    0.0, 0.0, 0.0, 0.0, 0.0]

    x4 = [-1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 
    -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 
    0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0,
     0.0, 0.0]

    x5 = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, 
    0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0, -1.0, 
    -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 
    0.0, 0.0, 0.0, 0.0]

    x6 = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, 
    1.0, -1.0, 1.0, -1.0, 1.0, -1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, -1.0, -1.0, 1.0, 
    1.0, -1.0, -1.0, 1.0, 1.0, -1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]


    level_x1 = []
    level_x2 = []
    level_x3 = []
    level_x4 = []
    level_x5 = []
    level_x6 = []

    print('请分别依次输入六个因子最高水平')
    x1_h, x2_h, x3_h, x4_h, x5_h, x6_h = map(float, input().split())
    print('''请分别依次输入六个因子最低水平
    和因子高水平顺序一致''')
    x1_l, x2_l, x3_l, x4_l, x5_l, x6_l = map(float, input().split())
    print('六个因子水平分别为：',x1_h ,x1_l ,x2_h ,x2_l, x3_h, x3_l, x4_h, x4_l, x5_h, x5_l, x6_h, x6_l)


    def get_real_level(x,x_h,x_l):
        level_x = []
        for i in x:
            if i == 0:
                level_x.append((x_h+x_l)/2)
            elif i == 1:
                level_x.append(x_h)
            else :
                level_x.append(x_l)
        return level_x

    level_x1 = get_real_level(x1,x1_h,x1_l)
    level_x2 = get_real_level(x2,x2_h,x2_l)
    level_x3 = get_real_level(x3,x3_h,x3_l)
    level_x4 = get_real_level(x4,x4_h,x4_l)
    level_x5 = get_real_level(x5,x5_h,x5_l)
    level_x6 = get_real_level(x6,x6_h,x6_l)

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
    new_sheet.write(0,4,'x5',style)
    new_sheet.write(0,5,'x6',style)
    new_sheet.write(0,6,'x1_real',style)
    new_sheet.write(0,7,'x2_real',style)
    new_sheet.write(0,8,'x3_real',style)
    new_sheet.write(0,9,'x4_real',style)
    new_sheet.write(0,10,'x5_real',style)
    new_sheet.write(0,11,'x6_real',style)
    new_sheet.write(0,12,'y1',style)
    new_sheet.write(0,13,'y2',style)

    for j in range(1,(len(x1)+1)):
        new_sheet.write(j,0,x1[j-1],style)
        new_sheet.write(j,1,x2[j-1],style)
        new_sheet.write(j,2,x3[j-1],style)
        new_sheet.write(j,3,x4[j-1],style)
        new_sheet.write(j,4,x5[j-1],style)
        new_sheet.write(j,5,x6[j-1],style)
        new_sheet.write(j,6,level_x1[j-1],style)
        new_sheet.write(j,7,level_x2[j-1],style)
        new_sheet.write(j,8,level_x3[j-1],style)
        new_sheet.write(j,9,level_x4[j-1],style)
        new_sheet.write(j,10,level_x5[j-1],style)
        new_sheet.write(j,11,level_x6[j-1],style)

    new_excel.save('D:/box_6factor.xlsx')
print('WELL 6')