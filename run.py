import xlrd
from xlrd import xldate_as_tuple
import datetime
from numpy import *

#load数据
data = xlrd.open_workbook(r'export effacement.xlsx')
excel = data.sheets()[0]
#读表头
table_array = {}
for col in range(excel.ncols):
    table_array.update({excel.cell_value(0, col): ''})
print(table_array)
# 获取数据
my_table =[]
my_table_all ={}
size = 28
usernum = excel.nrows//size #excel.nrows
for user_id in range(usernum ):
    rectange_MAE ,trapeze_MAE,kppv_MAE =[],[],[]
    puiss_real_kw = [] #取计算变量
    rectange_mae,rectange_pre= [],[]
    trapeze_mae = []
    if user_id < usernum-3:
        kppv_mae = []
    a =range(user_id*size+1, (user_id+1)*size+1)
    for rown in range(user_id*size+1, (user_id+1)*size+1):
        for col in range(excel.ncols):
            if excel.cell_value(0, col) == 'dh':  # 改变时间格式
                date = xldate_as_tuple(excel.cell(rown, col).value, 0)
                table_array[excel.cell_value(0, col)] = datetime.datetime(*date)
            else:
                table_array[excel.cell_value(0, col)] = excel.cell_value(rown, col)

            if excel.cell_value(0, col) == 'puiss_real_kw':
                puiss_real_kw.append(excel.cell_value(rown, col))
                if rown < user_id*size+4:#前三行用来预测
                    rectange_pre.append(excel.cell_value(rown, col))
                elif rown >= user_id*size+4 and rown < user_id*size+16:#4-15行 计算目标
                    rectange_mae.append(abs(excel.cell_value(rown, col)-mean(rectange_pre)))
                    #=($E$17-$E$4) / COUNT($E$4:$E$16)*COUNT($E$5: E5)+$E$4
                    trapeze_mae.append(abs(excel.cell_value(rown, col)-((excel.cell_value(user_id*size+16, col)-excel.cell_value(user_id*size+3, col))/((user_id*size+16)-(user_id*size+3))*(rown-(user_id*size+4)+1)+excel.cell_value(user_id*size+3, col))))
                    # kppv_mae.append()
                else:
                    pass
            if excel.cell_value(0, col) == 'p_install_kw' and user_id < usernum-3:
                if rown >= user_id*size+4 and rown < user_id*size+16:#4-15行 计算目标
                    ratio = []
                    for user_id_temp in range(user_id + 1, user_id + 4):
                        # print(user_id_temp)
                        # print("#######333",(user_id_temp-user_id) * size + rown,rown)
                        ratio.append(excel.cell_value((user_id_temp-user_id)* size + rown, 4) / excel.cell_value((user_id_temp-user_id) * size + rown, col))
                    kppv_mae.append(abs(excel.cell_value(rown, 4)-mean(ratio)*excel.cell_value(rown, col)))
                else:
                    pass
        # print(table_array)
        my_table.append(table_array)
    rectange_MAE.append(sum(rectange_mae))
    trapeze_MAE.append(sum(trapeze_mae))
    kppv_MAE.append(sum(kppv_mae))
print("rectange MAE = " ,mean(rectange_MAE))
print("trapeze MAE = " ,mean(trapeze_MAE))
print("kppv MAE = " ,mean(kppv_MAE))
