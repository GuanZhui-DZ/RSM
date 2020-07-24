import xlwt
def main_up():
    from up_factor_3 import up_b_factor_3
    from up_factor_4 import up_b_factor_4
    from up_factor_5 import up_b_factor_5
    from up_factor_2 import up_b_factor_2

    print("此程序仅可生成2-5因子二次回归通用旋转（RSM_UP）设计表")
    print('''请输入实验因子数''')
    factor_num = int(input())
    print(factor_num)

    if factor_num == 3:
        up_b_factor_3()

    if factor_num == 4:
        up_b_factor_4()

    if factor_num == 5:
        up_b_factor_5()

    if factor_num == 2:
        up_b_factor_2()


    input('d盘中查找excel表')
    print("done")