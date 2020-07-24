import xlwt
def main_box():
    from box_factor_3 import box_b_factor_3
    from box_factor_4 import box_b_factor_4
    from box_factor_5 import box_b_factor_5
    from box_factor_6 import box_b_factor_6

    print("此程序仅可生成3-6因子box_benhnken设计表")
    print('''请输入实验因子数''')
    factor_num = int(input())
    print(factor_num)

    if factor_num == 3:
        box_b_factor_3()

    if factor_num == 4:
        box_b_factor_4()

    if factor_num == 5:
        box_b_factor_5()

    if factor_num == 6:
        box_b_factor_6()


    input('d盘中查找excel表')
    print("done")