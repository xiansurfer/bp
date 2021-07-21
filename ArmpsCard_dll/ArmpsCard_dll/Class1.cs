using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using System.Data.OleDb;
using System.Collections;

using Matrix_Mul;
using NPOI.HSSF.UserModel;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace ArmpsCardProcessing
{
    public class ClassArmpsCardProcess
    {

        /// <summary>
        /// 用于矩阵的相关计算
        /// </summary>
        _Matrix_Calc matrix_Calc;

        _Matrix_Calc mat = new _Matrix_Calc();

        /// <summary>
        /// 导入的电流数据***格式为矩阵，每一列都是电流数据，各242个,多少列视数据而定
        /// </summary>
        _Matrix input_A;
        public _Matrix Input_A
        {
            set
            {
                this.input_A = value;
            }
            get
            {
                return input_A;
            }
        }

        /// <summary>
        /// 导出的电流卡片种类数据***格式为矩阵
        /// </summary>
        _Matrix output_A;
        public _Matrix Output_A
        {
            set
            {
                this.output_A = value;
            }
            get
            {
                return output_A;
            }
        }


        public _Matrix GetColumn(_Matrix data, int kk)
        {
            _Matrix p = new _Matrix(data.m, 1);
            p.init_matrix();
            for (int i = 0; i < data.m; i++)
            {
                p.write(i, 0, data.read(i, kk));
            }
            return p;
        }

        /// <summary>
        /// 归一化
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Normalize(_Matrix data)  //@_@ to do test
        {
            _Matrix dat = new _Matrix(data);
            double a = dat.read(0, 0);
            double b = dat.read(0, 1);
            double c = dat.read(0, 2);
            double d = dat.read(1, 0);

            for (int i = 0; i < dat.n; i++)
            {
                double min = 100000.0; double max = -100000.0;
                for (int j = 0; j < dat.m; j++)
                {
                    double s = dat.read(j, i);
                    if (s > max)
                    {
                        max = s;
                    }
                    else if (s < min)
                    {
                        min = s;
                    }
                }
                for (int j = 0; j < dat.m; j++)
                {
                    double s = dat.read(j, i);
                    s = s / max;
                    dat.write(j, i, s);
                }
            }
            return dat;
        }

        /// <summary>
        /// 按行错位，将第一行放到最后，其他行往前提一行
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix ReRank(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            for (int i = 0; i < dat.n; i++)
            {
                double s = dat.read(0, i);
                for (int j = 0; j < dat.m - 1; j++)
                {
                    dat.write(j, i, dat.read(j + 1, i));
                }
                dat.write(dat.m, i, s);
            }
            return dat;
        }

        /// <summary>
        /// 矩阵成员绝对值
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Element_ABS(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            for (int i = 0; i < dat.n; i++)
            {
                for (int j = 0; j < dat.m; j++)
                {
                    if (dat.read(j, i) < 0)
                    {
                        dat.write(j, i, -dat.read(j, i));
                    }
                }
            }
            return dat;
        }

        /// <summary>
        /// 矩阵成员求和
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Element_Sum(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            _Matrix sum = new _Matrix(data.n, 1);
            sum.init_matrix();
            for (int i = 0; i < dat.n; i++)
            {
                double s = 0;
                for (int j = 0; j < dat.m; j++)
                {
                    s += dat.read(j, i);
                }
                sum.write(i, 0, s);
            }
            return sum;
        }

        /// <summary>
        /// 矩阵成员前后半程分别求取最大值（用于求取最大值前后的极大值）
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Element_Max(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            _Matrix Max = new _Matrix(data.m, 2);
            double a = dat.read(0, 0);
            double b = dat.read(0, 1);
            double c = dat.read(1, 0);
            Max.init_matrix();
            for (int i = 0; i < dat.m; i++)
            {
                double max = -100000.0;
                for (int j = 0; j < dat.n / 2; j++)
                {
                    double s = dat.read(i, j);
                    if (s > max)
                    {
                        max = s;
                    }
                }
                Max.write(i, 0, max);

                max = -100000.0;
                for (int j = dat.n / 2 + 1; j < dat.n; j++)
                {
                    double s = dat.read(i, j);
                    if (s > max)
                    {
                        max = s;
                    }
                }
                Max.write(i, 1, max);
            }
            return Max;
        }

        /// <summary>
        /// 矩阵成员前后半程分别求取最小值（用于求取最大值前后的极小值）
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Element_Min(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            _Matrix Min = new _Matrix(data.m, 2);
            double a = dat.read(0, 0);
            double b = dat.read(0, 1);
            double c = dat.read(1, 0);
            Min.init_matrix();
            for (int i = 0; i < dat.m; i++)
            {
                double min = 100000.0;
                for (int j = 0; j < dat.n / 2; j++)
                {
                    double s = dat.read(i, j);
                    if (s < min)
                    {
                        min = s;
                    }
                }
                Min.write(i, 0, min);

                min = 100000.0;
                for (int j = dat.n / 2 + 1; j < dat.n; j++)
                {
                    double s = dat.read(i, j);
                    if (s < min)
                    {
                        min = s;
                    }
                }
                Min.write(i, 1, min);
            }
            return Min;
        }


        /// <summary>
        /// 矩阵成员最大值前后的六个元素，共7个元素
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Element_MaxSeries(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            _Matrix MaxSeries = new _Matrix(data.n, 7);
            MaxSeries.init_matrix();
            for (int i = 0; i < dat.n; i++)
            {
                double max = -100000.0;
                int maxNum = 0;
                for (int j = 0; j < dat.m; j++)
                {
                    double s = dat.read(j, i);
                    if (s > max)
                    {
                        max = s;
                        maxNum = j;

                        MaxSeries.write(i, 0, dat.read(j - 3 <= 0 ? 0 : (j - 3), i));//存在疑似负值，先定为0
                        MaxSeries.write(i, 1, dat.read(j - 2 <= 0 ? 0 : (j - 2), i));
                        MaxSeries.write(i, 2, dat.read(j - 1 <= 0 ? 0 : (j - 1), i));
                        MaxSeries.write(i, 4, dat.read(j + 1 >= dat.m ? dat.m : (j + 1), i));
                        MaxSeries.write(i, 5, dat.read(j + 2 >= dat.m ? dat.m : (j + 2), i));
                        MaxSeries.write(i, 6, dat.read(j + 3 >= dat.m ? dat.m : (j + 3), i));
                    }
                }
                MaxSeries.write(i, 3, max);
            }
            return MaxSeries;
        }


        /// <summary>
        /// 矩阵成员对应相除,用电流总和除以电流波动总和
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Division(_Matrix data1, _Matrix data2)
        {
            _Matrix dat1 = new _Matrix(data1);
            _Matrix dat2 = new _Matrix(data2);
            _Matrix division = new _Matrix(data1.m, 1);
            division.init_matrix();
            for (int i = 0; i < dat1.m; i++)
            {
                for (int j = 0; j < dat1.n; j++)
                {
                    division.write(i, j, dat1.read(i, j) / data2.read(i, j));
                }
            }
            return division;
        }


        /// <summary>
        /// 求矩阵中的零值和非零值个数，其中第一维为0值个数，第二维是非零值个数
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix CountZero(_Matrix data)
        {
            _Matrix dat = new _Matrix(data);
            _Matrix countzero = new _Matrix(data.n, 2);
            countzero.init_matrix();
            for (int i = 0; i < dat.n; i++)
            {
                int count = 0;
                for (int j = 0; j < dat.m; j++)
                {
                    double s = dat.read(j, i);
                    if (Math.Abs(s) < 0.00001)
                    {
                        count++;
                    }
                }
                countzero.write(i, 0, count);
                countzero.write(i, 1, 241 - count);
            }
            return countzero;
        }


        public _Matrix OutputProcessing(_Matrix Output_A)
        {
            _Matrix dat = new _Matrix(Output_A);
            int max = -100;
            for (int i = 0; i < dat.n; i++)
            {
                int s = (int)dat.read(0, i);
                if (s > max)
                {
                    max = s;
                }
            }
            _Matrix Output_A_Norm = new _Matrix(max + 1, dat.n);
            Output_A_Norm.init_matrix();
            for (int i = 0; i < dat.n; i++)
            {
                int s = (int)dat.read(0, i);
                Output_A_Norm.write(s, i, 1);//是第几种工况就在第几个元素处写成1,0预留为未定工况
            }
            double a = Output_A_Norm.read(0, 0);
            double b = Output_A_Norm.read(0, 1);
            double c = Output_A_Norm.read(0, 2);
            double d = Output_A_Norm.read(0, 3);
            double e = Output_A_Norm.read(1, 0);//1
            double g = Output_A_Norm.read(1, 1);//1
            double h = Output_A_Norm.read(2, 1);
            double f = Output_A_Norm.read(2, 0);
            return Output_A_Norm;

        }


        public _Matrix DataPreprocessingAndPrint(_Matrix Input_A, _Matrix Output_A)
        {
            _Matrix_Calc matrix_Calc = new _Matrix_Calc();
            _Matrix Input_A_Norm = Normalize(Input_A);//数据标准化
            _Matrix SumInput_A_Norm = Element_Sum(Input_A_Norm);//电流数据求和，反映电流的总情况，特征数据1
            _Matrix Input_A_Norm_ReRank = ReRank(Input_A_Norm);//数据错位方便进行做差求差的绝对值

            _Matrix DetInput_A_Norm = matrix_Calc.subtracts(Input_A_Norm, Input_A_Norm_ReRank);//错位做差
            _Matrix AbsDetInput_A_Norm = Element_ABS(DetInput_A_Norm);//利用错位数据做的差求取绝对值
            _Matrix SumAbsDetInput_A_Norm = Element_Sum(AbsDetInput_A_Norm);//单列电流波动值的和，特征数据2
            _Matrix MaxSeriesAbsDetInput_A_Norm = Element_MaxSeries(AbsDetInput_A_Norm);//求取电流波动情况绝对值的最大值（前后共7个，其中第4个是最大值），最大值为特征数据3
            _Matrix MinOfMaxSeries = Element_Min(MaxSeriesAbsDetInput_A_Norm);//电流波动最大前后的极小电流波动值，共两个数，特征数据4,5
            _Matrix MaxOfMaxSeries = Element_Max(MaxSeriesAbsDetInput_A_Norm);//电流波动最大前后的极大电流波动值，共两个数，特征数据6,7
            /*********此处调整参数顺序，用绝对值最大除以总值*********/
            _Matrix DivisionOfSumandSumAbs = Division(SumAbsDetInput_A_Norm, SumInput_A_Norm);//电流总和除以电流波动值的和，特征数据8
            _Matrix CountZero_Input_A_Norm = CountZero(Input_A_Norm);//电流中0值和非零值的个数，特征数据9,10


            _Matrix MixtureMatrix = new _Matrix(Input_A_Norm.n, 11);
            MixtureMatrix.init_matrix();
            for (int i = 0; i < Input_A_Norm.n; i++)
            {
                MixtureMatrix.write(i, 0, SumInput_A_Norm.read(i, 0) / 241);//此处修改为平均值
                MixtureMatrix.write(i, 1, SumAbsDetInput_A_Norm.read(i, 0));
                MixtureMatrix.write(i, 2, MaxSeriesAbsDetInput_A_Norm.read(i, 3));
                MixtureMatrix.write(i, 3, MinOfMaxSeries.read(i, 0));
                MixtureMatrix.write(i, 4, MinOfMaxSeries.read(i, 1));

                /*****修改后的最大值定义******/
                MixtureMatrix.write(i, 5, MaxSeriesAbsDetInput_A_Norm.read(i, 3) - MinOfMaxSeries.read(i, 0));
                MixtureMatrix.write(i, 6, MaxSeriesAbsDetInput_A_Norm.read(i, 3) - MinOfMaxSeries.read(i, 1));

                //MixtureMatrix.write(i, 5, MaxOfMaxSeries.read(i, 0));
                //MixtureMatrix.write(i, 6, MaxOfMaxSeries.read(i, 1));
                MixtureMatrix.write(i, 7, DivisionOfSumandSumAbs.read(i, 0) * 241);//修改后的求除法，用波动求和除以电流平均值（原为除以电流总值）
                MixtureMatrix.write(i, 8, CountZero_Input_A_Norm.read(i, 0));
                MixtureMatrix.write(i, 9, CountZero_Input_A_Norm.read(i, 1));
                MixtureMatrix.write(i, 10, Output_A.read(0, i));

            }
            return MixtureMatrix;


        }
    }
}
