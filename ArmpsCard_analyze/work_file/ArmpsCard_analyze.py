import os
import clr
import numpy as np
import pandas as pd
import xlwt
import math
clr.FindAssembly('ArmpsCardProcessing.dll')
clr.AddReference('ArmpsCardProcessing')
clr.FindAssembly('Matrix_Mul.dll')
clr.AddReference('Matrix_Mul')
clr.FindAssembly('BP.dll')
clr.AddReference('BP')
import ArmpsCardProcessing
import Matrix_Mul
import BPNETSerial

"""
将数组data_arr中的数据写入初始化好的Matrix_init对象
返回一个被写入值的Matrix对象
"""
def Matrix_wirite(Matrix_init,data_arr):
    rows = data_arr.shape[0]
    columns = data_arr.shape[1]
    j = 0
    for i in range(rows):
        for k in range(columns):
            Matrix_init.arr[j] = data_arr[i][k]
            j += 1

    return Matrix_init

"""
用于快速查看Matrix对象中的矩阵值
"""
def Matrix_check(Matrix):
    data = []
    for i in range(Matrix.m):
        row = []
        for j in range(Matrix.n):
            row.append(Matrix.read(i, j))
        data.append(row)
    return pd.DataFrame(data)

"""
中间处理结果输出写入Excel,并返回trainmatrix数组
"""
def output_Excel_DataProcessing(Input_A,MixtureMatrix,OutputMatrix,WorkingCondition):
    workbooks = xlwt.Workbook()
    worksheet = workbooks.add_sheet('sheet1')
    for i in range(1,Input_A.n+1):
        for j in range(1,12):
            worksheet.write(i,j - 1,MixtureMatrix.read(i-1,j-1))
        for k in range(1,WorkingCondition+1):
            worksheet.write(i, k -1 + 11, OutputMatrix.read(k-1,i-1))

    for j in range(1,12):
        if j < 11:
            worksheet.write(0,j-1,'特征值%i'%(j))
        else:
            worksheet.write(0,j-1,'工况情况')

    for k in range(1,WorkingCondition+1):
        worksheet.write(0, k - 1 + 11,'工况矩阵%i'%(k-1))
    workbooks.save('NormMatrix.xlsx')
    df = pd.read_excel('NormMatrix.xlsx')
    return df.values

def current_data_analyze(path):
    # 读取excel文件
    df_base = pd.read_excel(path,header=None)
    data = df_base.values
    data_input = data[1:,:]
    data_output = data[0,:].reshape(1,data_input.shape[1])
    # 获取文件列数
    columns = df_base.shape[1]

    """
    创建ClassArmpsCardProcess的实例对象CACP，并给CACP的Input_A和Output_A赋值
    Input_A 赋data[1:,:]
    Output_A 赋data[0,:]
    """
    CACP = ArmpsCardProcessing.ClassArmpsCardProcess()
    CACP.Input_A = Matrix_Mul._Matrix(241,columns)
    CACP.Input_A.init_matrix()
    CACP.Output_A = Matrix_Mul._Matrix(1,columns)
    CACP.Output_A.init_matrix()

    """ ******Input中存储输入的电流数据******* """
    CACP.Input_A = Matrix_wirite(CACP.Input_A,data_input)
    # df_CACP_Input_A = Matrix_check(CACP.Input_A)

    """ *****Output中存储输出结果**** """
    CACP.Output_A = Matrix_wirite(CACP.Output_A,data_output)
    # df_CACP_Output_A = Matrix_check(CACP.Output_A)

    """ 处理数据形成特征矩阵，附带一列结果向量 """
    MixtureMatrix = CACP.DataPreprocessingAndPrint(CACP.Input_A, CACP.Output_A)
    # df_MixtureMatrix = Matrix_check(MixtureMatrix)

    """ 处理完成后形成基于结果的数据矩阵 """
    OutputMatrix = CACP.OutputProcessing(CACP.Output_A)
    # df_OutputMatrix = Matrix_check(OutputMatrix)

    dat2 = output_Excel_DataProcessing(CACP.Input_A, MixtureMatrix, OutputMatrix, CACP.WorkingCondition)
    dat2_Rows_Count = dat2.shape[0]
    dat2_Columns_Count = dat2.shape[1]

    InputLayerNum = 10
    HiddenLayerNum = 15
    OutputLayerNum = 10

    """ 所有行，以及前10列为电流的特征数据 """
    Matrix_Feature_Pred = Matrix_Mul._Matrix(dat2_Rows_Count, InputLayerNum)

    """ 所有行，除了前11列以外的数据为电流的输出矩阵 """
    Matrix_Output_Pred = Matrix_Mul._Matrix(dat2_Rows_Count,dat2_Columns_Count - InputLayerNum - 2)

    Matrix_Feature_Pred.init_matrix()
    Matrix_Output_Pred.init_matrix()

    """ 20210716扩充工况类型的变量修复 """
    OutputLayerNum = Matrix_Output_Pred.n

    data_Feature_Pred = dat2[:,:InputLayerNum]
    data_Output_Pred = dat2[:,InputLayerNum+2:dat2_Columns_Count]

    Matrix_Feature_Pred = Matrix_wirite(Matrix_Feature_Pred,data_Feature_Pred)
    Matrix_Output_Pred = Matrix_wirite(Matrix_Output_Pred,data_Output_Pred)

    """ 初始化bp神经网络，参数分别为输入层节点数，隐层节点，输出层节点，数据样本数，数据维度（对应输入层节点数），迭代步数，学习率"""
    bp = BPNETSerial.BP(InputLayerNum, HiddenLayerNum, OutputLayerNum, dat2_Rows_Count, OutputLayerNum, 2000, 0.2)
    mat = Matrix_Mul._Matrix_Calc()
    bp.b2.init_matrix()
    bp.b1.init_matrix()
    bp.w1.init_matrix()
    bp.w2.init_matrix()

    """ 读取各权重系数，并写入Matrix """
    arr_b2 = pd.read_excel('saveWB.xls',sheet_name='b2',header=None).values
    bp.b2 = Matrix_wirite(bp.b2,arr_b2)
    arr_w2 = pd.read_excel('saveWB.xls',sheet_name='w2',header=None).values
    bp.w2 = Matrix_wirite(bp.w2,arr_w2)
    arr_b1 = pd.read_excel('saveWB.xls', sheet_name='b1', header=None).values
    bp.b1 = Matrix_wirite(bp.b1, arr_b1)
    arr_w1 = pd.read_excel('saveWB.xls', sheet_name='w1', header=None).values
    bp.w1 = Matrix_wirite(bp.w1, arr_w1)

    bp.Input_test = mat.transposs(Matrix_Feature_Pred)
    bp.Output_test = mat.transposs(Matrix_Output_Pred)
    bp.testBP(bp.Input_test, bp.Output_test, dat2_Rows_Count)
    bp.ConvNorm()
    bp.CalcAccuracy()

    bp.GetComprehensiveEvaluationMatrix(4)
    bp.comprehensiveEvaluation = [0.3009, 0.2201, 0.2696, 0.1793]
    bp.ComEvaluation(bp.comprehensiveEvaluation)

    return bp.output_test_Norm, bp.fore_test

def result_process(result_Matrix):
    for i in range(result_Matrix.m):
        sum = 0
        for j in range(result_Matrix.n):
            if result_Matrix.read(i,j) < 0:
                result_Matrix.write(i,j,0)
            sum += result_Matrix.read(i, j)
        for j in range(result_Matrix.n):
            result_Matrix.write(i, j, result_Matrix.read(i, j) / sum)

    return result_Matrix

def output_Excel(output_test_Norm,fore_test):
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('sheet1')
    fore_test = result_process(fore_test)
    for i in range(1,fore_test.m+1):
        max = -100
        flag = fore_test.m*[0]
        for j in range(1,fore_test.n + 1):
            worksheet.write(i-1,j-1,fore_test.read(i - 1, j - 1))
            if fore_test.read(i - 1, j - 1) > max:
                max = fore_test.read(i-1, j-1)
                flag[i-1] = j

        if (abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 1)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 2)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 3)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 4)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 5)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 6)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 7)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 8)) < 0.00001 and
                abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 9)) < 0.00001):
            flag[i - 1] = 11

        worksheet.write(i-1,fore_test.n,flag[i - 1])
        workbook.save('result.xlsx')

output_test_Norm,fore_test = current_data_analyze('ArmpsCardData12工况.xls')
output_Excel(output_test_Norm,fore_test)
