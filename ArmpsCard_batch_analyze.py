import os
import clr
import pandas
import numpy
import pandas as pd

clr.FindAssembly('ArmpsCardProcessing.dll')
clr.AddReference('ArmpsCardProcessing')
clr.FindAssembly('Matrix_Mul.dll')
clr.AddReference('Matrix_Mul')

import ArmpsCardProcessing
import Matrix_Mul

def Form1(path):
    # 读取excel文件
    df_base = pd.read_excel(path,header=None)
    data = df_base.values
    # 获取文件列数
    columns = df_base.shape[1]

    """
    创建ClassArmpsCardProcess的实例对象CACP，并给CACP的Input_A和Output_A赋值
    Input_A 赋[1:241,:]
    Output_A 赋[0,:]
    """
    CACP = ArmpsCardProcessing.ClassArmpsCardProcess()
    CACP.Input_A = Matrix_Mul._Matrix(241,columns)
    CACP.Input_A.init_matrix()
    j = 0
    for i in range(1,242):
        for k in range(data.shape[1]):
            CACP.Input_A.arr[j] = data[i][k]
            j += 1






Form1('ArmpsCardData12工况.xls')