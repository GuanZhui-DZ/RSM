import xlwt
from box_main import main_box
from up_main import main_up

print("此程序仅可创建二次回归通用旋转设计表和BOX_BENHNKEN表")
print("输入1开始创建二次回归通用旋转设计表")
print("输入2开始创建BOX_BENHNKEN表")
num = int(input())

if num == 1:
    main_up()
if num == 2:
    main_box()

print("done")