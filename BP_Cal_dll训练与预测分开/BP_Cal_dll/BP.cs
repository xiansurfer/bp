using Matrix_Mul;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;



namespace BPNETSerial
{
    public class BP
    {

        /// <summary>
        /// 判断是否训练过网络
        /// </summary>
        public Boolean IsTrained;
        /// <summary>
        /// 用于矩阵的相关计算
        /// </summary>
        public _Matrix_Calc matrix_Calc;

        /// <summary>
        /// 输入层节点数
        /// </summary>
        public int innum;

        /// <summary>
        /// 测试数据组数
        /// </summary>
        public int train_num;
        /// <summary>
        /// 训练数据组数
        /// </summary>

        public int test_num;
        public int Test_num
        {
            get
            {
                return test_num;
            }
            set
            {
                test_num = value;
            }
        }
        /// <summary>
        /// 测试数据维度;
        /// </summary>
        public int sampdim;

        /// <summary>
        /// 隐藏层节点数
        /// </summary>
        public int midnum;

        /// <summary>
        /// 输出层节点数
        /// </summary>
        public int outnum;

        /// <summary>
        /// 迭代次数
        /// </summary>
        public int iteration;

        /// <summary>
        /// 输入层与隐藏层间的权值
        /// </summary>
        public _Matrix w1;

        /// <summary>
        /// 输入层与隐藏层间的阀值
        /// </summary>
        public _Matrix b1;

        /// <summary>
        /// 输出层与隐藏层间的权值
        /// </summary>
        public _Matrix w2;

        /// <summary>
        /// 输出层与隐藏层间的阀值
        /// </summary>
        public _Matrix b2;

        /// <summary>
        /// 保存w1的值
        /// </summary>
        public _Matrix w1_1;

        /// <summary>
        /// 保存w2的值
        /// </summary>
        public _Matrix w2_1;

        /// <summary>
        /// 用于综合评价的矩阵（基于bp神经网络测试结果）
        /// </summary>
        public _Matrix comprehesiveEvaluationMatrix;
        /// <summary>
        /// 综合评价结果输出矩阵（基于bp神经网络测试结果）
        /// </summary>
        public _Matrix comprehensiveEvaluationResultMatrix;
        /// <summary>
        /// 保存b1的值
        /// </summary>
        public _Matrix b1_1;

        /// <summary>
        /// 保存b2的值
        /// </summary>
        public _Matrix b2_1;

        /// <summary>
        /// 学习率
        /// </summary>
        public double xite;

        /// <summary>
        /// 误差
        /// </summary>
        public double error;

        public double[] comprehensiveEvaluation;

        public double accu_average;
        /// <summary>
        /// 误差率
        /// </summary>
        public double[] accuracy;
        /// <summary>
        /// 训练输入数据
        /// </summary>
        public _Matrix input_train;

        public _Matrix Input_train
        {
            get
            {
                return input_train;
            }
            set
            {
                this.input_train = value;
            }
        }
        /// <summary>
        /// 训练输出数据
        /// </summary>
        public _Matrix output_train;

        public _Matrix Output_train
        {
            get
            {
                return output_train;
            }
            set
            {
                this.output_train = value;
            }
        }
        /// <summary>
        /// 归一化后的训练输入数据
        /// </summary>
        public _Matrix input_train_Norm;

        /// <summary>
        /// 归一化后的训练输出数据
        /// </summary>
        public _Matrix output_train_Norm;

        /// <summary>
        /// 测试输入数据
        /// </summary>
        public _Matrix input_test;

        public _Matrix Input_test
        {
            get
            {
                return input_test;
            }
            set
            {
                this.input_test = value;
            }
        }

        /// <summary>
        /// 预期输出数据(归一化前)
        /// </summary>
        public _Matrix fore_test;

        /// <summary>
        /// 预期输出数据(归一化后）
        /// </summary>
        public _Matrix fore;

        /// <summary>
        /// 测试输出数据
        /// </summary>
        public _Matrix output_test;

        public _Matrix Output_test
        {
            get
            {
                return output_test;
            }
            set
            {
                this.output_test = value;
            }
        }

        /// <summary>
        /// 误差矩阵
        /// </summary>
        public _Matrix error_test;

        /// <summary>
        /// 归一化后的测试输入数据
        /// </summary>
        public _Matrix input_test_Norm;

        /// <summary>
        /// 归一化后的测试输出数据
        /// </summary>
        public _Matrix output_test_Norm;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="innum"></param>
        /// <param name="midnum"></param>
        /// <param name="outnum"></param>
        /// <param name="num"></param>
        /// <param name="sampDim"></param>
        /// <param name="input_train"></param>
        /// <param name="output_train"></param>
        /// <param name="xite"></param>
        public BP(int innum, int midnum, int outnum, int train_num, int sampDim, int iteration, double xite)
        {
            this.innum = innum;
            this.midnum = midnum;
            this.outnum = outnum;
            this.iteration = iteration;
            matrix_Calc = new _Matrix_Calc();
            this.train_num = train_num;
            this.sampdim = sampDim;
            this.xite = xite;
            this.input_train = new _Matrix(train_num, sampDim);
            input_train.init_matrix();
            this.output_train = new _Matrix(train_num, outnum);
            output_train.init_matrix();

            //初始化w1,w2,b1,b2;
            w1 = InitWB(midnum, innum);
            w2 = InitWB(midnum, outnum);
            b1 = InitWB(midnum, 1);
            b2 = InitWB(outnum, 1);
            w1_1 = new _Matrix(w1);
            b1_1 = new _Matrix(b1);
            w2_1 = new _Matrix(w2);
            b2_1 = new _Matrix(b2);
        }

        /// <summary>
        /// 应用与BP训练时的计算，矩阵的每一个值乘上学习率
        /// </summary>
        /// <param name="data"></param>
        /// <param name="xite"></param>
        /// <returns></returns>
        public _Matrix AddStudyRate(_Matrix data, double xite)
        {
            for (int i = 0; i < data.m; i++)
            {
                for (int j = 0; j < data.n; j++)
                {
                    data.write(i, j, data.read(i, j) * xite);
                }
            }
            return data;

        }

        /// <summary>
        /// 归一化
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public _Matrix Normalize(_Matrix data)  //@_@ to do test
        {
            _Matrix dat = new _Matrix(data);
            for (int i = 0; i < dat.m; i++)
            {
                double min = 100000.0; double max = -100000.0;
                for (int j = 0; j < dat.n; j++)
                {
                    double s = dat.read(i, j);
                    if (s > max)
                    {
                        max = s;
                    }
                    else if (s < min)
                    {
                        min = s;
                    }


                }
                for (int j = 0; j < dat.n; j++)
                {
                    double s = dat.read(i, j);
                    s = (s - min) / (max - min);
                    dat.write(i, j, s);
                }
            }

            return dat;
        }

        /// <summary>
        /// 初始化w1,w2,b1,b2
        /// </summary>
        /// <param name="m"></param>
        /// <param name="n"></param>
        /// <returns></returns>
        public _Matrix InitWB(int m, int n)
        {
            _Matrix mat = new _Matrix(m, n);
            mat.init_matrix();
            Random rand = new Random();
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    double s;
                    s = (rand.NextDouble() - 0.5) * 2;
                    mat.write(i, j, s);
                }
            }
            return mat;

        }

        /// <summary>
        /// 获取矩阵的某一行
        /// </summary>
        /// <param name="data"></param>
        /// <param name="kk"></param>
        /// <returns></returns>
        public _Matrix GetRow(_Matrix data, int kk)
        {
            _Matrix p = new _Matrix(1, data.n);
            p.init_matrix();
            for (int i = 0; i < data.n; i++)
            {
                p.write(0, i, data.read(kk, i));
            }
            return p;
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

        public double sumsqr(_Matrix data)
        {
            double s = 0.0;
            for (int i = 0; i < data.m; i++)
            {
                for (int j = 0; j < data.n; j++)
                {
                    s += data.read(i, j) * data.read(i, j);
                }
            }
            return s;
        }


        /// <summary>
        /// 训练网络
        /// </summary>
        /// <param name="input_train">输入矩阵（可以是未归一化的）</param>
        /// <param name="output_train">期望的输出矩阵（可以是未归一化的）</param>
        public void trainBP(_Matrix input_train, _Matrix output_train)
        {
            this.input_train = input_train;
            this.output_train = output_train;
            this.input_train_Norm = Normalize(input_train);
            this.output_train_Norm = Normalize(output_train);

            for (int ii = 0; ii < iteration; ii++)
            {
                for (int i = 0; i < train_num; i++)
                {
                    var x = GetColumn(input_train_Norm, i);

                    _Matrix I = new _Matrix(1, midnum);
                    I.init_matrix();
                    _Matrix lout = new _Matrix(1, midnum);
                    lout.init_matrix();
                    for (int j = 0; j < midnum; j++)
                    {
                        _Matrix t = new _Matrix(1, 1);
                        t.init_matrix();

                        _Matrix aaa = GetColumn(input_train_Norm, i);
                        _Matrix tt = matrix_Calc.transposs(aaa);
                        _Matrix ttt = matrix_Calc.transposs(GetRow(w1, j));
                        t = matrix_Calc.multiplys(tt, ttt);
                        I.write(0, j, t.read(0, 0) + b1.read(j, 0));
                        double s = 1 / (1 + Math.Exp(-I.read(0, j)));
                        lout.write(0, j, s);

                    }

                    _Matrix yn;
                    _Matrix y = matrix_Calc.transposs(w2);
                    _Matrix yy = matrix_Calc.transposs(lout);
                    _Matrix yyy = matrix_Calc.multiplys(y, yy);
                    yn = matrix_Calc.adds(yyy, b2);
                    _Matrix e;
                    e = GetColumn(output_train_Norm, i);
                    e = matrix_Calc.subtracts(e, yn);
                    _Matrix dw2;
                    dw2 = matrix_Calc.multiplys(e, lout);
                    _Matrix db2 = matrix_Calc.transposs(e);
                    _Matrix dw1 = new _Matrix(innum, midnum);
                    dw1.init_matrix();
                    _Matrix db1 = new _Matrix(1, midnum);
                    db1.init_matrix();
                    double[] FI = new double[midnum];

                    for (int j = 0; j < midnum; j++)
                    {
                        double S = 1 / (1 + Math.Exp(-I.read(0, j)));
                        FI[j] = S;
                    }
                    for (int k = 0; k < innum; k++)
                    {
                        for (int j = 0; j < midnum; j++)
                        {
                            double s = 0.0;
                            for (int tt = 0; tt < outnum; tt++)
                            {
                                s += e.arr[tt] * w2.read(j, tt);
                            }
                            dw1.write(k, j, FI[j] * x.read(k, 0) * s);
                            db1.write(j, 1, FI[j] * s);
                        }
                    }
                    _Matrix sw1 = matrix_Calc.transposs(dw1);
                    _Matrix sb1 = matrix_Calc.transposs(db1);
                    _Matrix sw2 = matrix_Calc.transposs(dw2);
                    _Matrix sb2 = matrix_Calc.transposs(db2);
                    w1 = matrix_Calc.adds(w1_1, AddStudyRate(sw1, xite));
                    _Matrix aaaa = AddStudyRate(sb1, xite);
                    b1 = matrix_Calc.adds(b1_1, aaaa);
                    w2 = matrix_Calc.adds(w2_1, AddStudyRate(sw2, xite));
                    b2 = matrix_Calc.adds(b2_1, AddStudyRate(sb2, xite));
                    w1_1 = new _Matrix(w1);
                    b1_1 = new _Matrix(b1);
                    w2_1 = new _Matrix(w2);
                    b2_1 = new _Matrix(b2);

                }


            }

        }

        /// <summary>
        /// 将测试数据代入进行测试
        /// </summary>
        /// <param name="input_test">测试组的输入数据</param>
        /// <param name="output_test">测试组的预期输出数据</param>
        /// <param name="test_num">测试组的组数</param>
        public void testBP(_Matrix input_test, _Matrix output_test, int test_num)
        {
            this.input_test = input_test;
            this.output_test = output_test;
            this.input_test_Norm = Normalize(input_test);
            this.output_test_Norm = Normalize(output_test);
            fore_test = new _Matrix(output_test.m, output_test.n);
            fore_test.init_matrix();
            error_test = new _Matrix(output_test.m, output_test.n);
            error_test.init_matrix();
            this.test_num = test_num;
            for (int i = 0; i < test_num; i++)
            {
                double[] I = new double[midnum];
                _Matrix lout = new _Matrix(1, midnum);
                lout.init_matrix();
                for (int j = 0; j < midnum; j++)
                {

                    _Matrix s = GetColumn(input_test_Norm, i);
                    s = matrix_Calc.transposs(s);
                    _Matrix ss = GetRow(w1, j);
                    ss = matrix_Calc.transposs(ss);
                    _Matrix sss = matrix_Calc.multiplys(s, ss);
                    I[j] = sss.arr[0] + b1.read(j, 0);

                    lout.write(0, j, 1 / (1 + Math.Exp(-I[j])));


                }
                _Matrix t = matrix_Calc.transposs(w2);
                _Matrix tt = matrix_Calc.transposs(lout);
                _Matrix ttt = matrix_Calc.adds(matrix_Calc.multiplys(t, tt), b2);
                for (int j = 0; j < fore_test.m; j++)
                {
                    fore_test.write(j, i, ttt.read(j, 0));
                }

                ////20210713增加用于识别启停井当天，若当天0值数据占比超过0.3，判定为启停井当天
                if (input_test_Norm.read(8, i) + input_test_Norm.read(9, i) == 0)
                {
                    for (int j = 0; j < fore_test.m; j++)
                    {
                        fore_test.write(j, i, 1);
                    }
                }
                ////20210713增加部分结束

            }
            error_test = matrix_Calc.subtracts(fore_test, output_test_Norm);
            error = sumsqr(error_test);
            //Console.WriteLine(error);

        }


        /// <summary>
        /// 获得综合评价矩阵
        /// </summary>
        /// <param name="num">综合评价矩阵的维度</param>
        public void GetComprehensiveEvaluationMatrix(int num)
        {
            if (output_test_Norm.arr == null)
            {
                return;
            }
            output_test_Norm = matrix_Calc.transposs(output_test_Norm);
            fore_test = matrix_Calc.transposs(fore_test);
            comprehesiveEvaluationMatrix = new _Matrix(output_test_Norm.m, 2 * num);
            comprehesiveEvaluationMatrix.init_matrix();
            for (int i = 0; i < comprehesiveEvaluationMatrix.m; i++)
            {
                comprehesiveEvaluationMatrix.write(i, 0, output_test_Norm.read(i, 0));
                //double s = 0.0;
                //s = output_test_Norm.read(i,1)+output_test_Norm.read(i,2)+output_test_Norm.read(i,3)+output_test_Norm.read(i,4)+output_test_Norm.read(i,5);
                //s = s / 5;
                //comprehesiveEvaluationMatrix.write(i, 1, s);
                //comprehesiveEvaluationMatrix.write(i, 2, output_test_Norm.read(i, 6));
                //comprehesiveEvaluationMatrix.write(i, 3, output_test_Norm.read(i, 7));


                comprehesiveEvaluationMatrix.write(i, 1, output_test_Norm.read(i, 1));
                comprehesiveEvaluationMatrix.write(i, 2, output_test_Norm.read(i, 2));
                comprehesiveEvaluationMatrix.write(i, 3, output_test_Norm.read(i, 3));
                comprehesiveEvaluationMatrix.write(i, 4, fore_test.read(i, 0));
                comprehesiveEvaluationMatrix.write(i, 5, fore_test.read(i, 1));
                comprehesiveEvaluationMatrix.write(i, 6, fore_test.read(i, 2));
                comprehesiveEvaluationMatrix.write(i, 7, fore_test.read(i, 3));

                //comprehesiveEvaluationMatrix.write(i, 4, fore_test.read(i, 0));

                //s = fore_test.read(i, 1) + fore_test.read(i, 2) + fore_test.read(i, 3) + fore_test.read(i, 4) + fore_test.read(i, 5);
                //s = s / 5;
                //comprehesiveEvaluationMatrix.write(i, 5, s);
                //comprehesiveEvaluationMatrix.write(i, 6, fore_test.read(i, 6));
                //comprehesiveEvaluationMatrix.write(i, 7, fore_test.read(i, 7));
            }
        }
        /// <summary>
        /// 进行综合评价,获得综合评价后的结果矩阵
        /// </summary>
        /// <param name="ComEval">各维度权值</param>
        /// <param name="data">评价矩阵</param>
        public void ComEvaluation(double[] ComEval)
        {
            comprehensiveEvaluationResultMatrix = new _Matrix(comprehesiveEvaluationMatrix.m, 3);
            comprehensiveEvaluationResultMatrix.init_matrix();
            _Matrix data = new _Matrix(comprehesiveEvaluationMatrix);
            for (int i = 0; i < data.m; i++)
            {
                double s = 0.0;
                for (int j = 0; j < data.n / 2; j++)
                {
                    s += ComEval[j] * data.read(i, j);
                }
                comprehensiveEvaluationResultMatrix.write(i, 0, i + 1);
                comprehensiveEvaluationResultMatrix.write(i, 1, s);
                s = 0.0;
                for (int j = data.n / 2; j < data.n; j++)
                {
                    s += ComEval[j - data.n / 2] * data.read(i, j);
                }

                comprehensiveEvaluationResultMatrix.write(i, 2, s);
            }
        }



        /// <summary>
        /// ComEvaResult矩阵写入EXCEL
        /// </summary>
        public void ComEvaResult_Excel()
        {
            if (comprehensiveEvaluationResultMatrix.arr == null)
            {
                return;
            }
            var excelApp = new Microsoft.Office.Interop.Excel.Application();


            Workbooks workbooks = excelApp.Workbooks;
            Workbook workBook = workbooks.Add(Type.Missing);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1

            for (int i = 1; i <= comprehensiveEvaluationResultMatrix.m; i++)
            {
                for (int j = 1; j <= comprehensiveEvaluationResultMatrix.n; j++)
                {
                    workSheet.Cells[i, j] = comprehensiveEvaluationResultMatrix.read(i - 1, j - 1);
                }
            }

            workBook.SaveAs(@"f:\BPSeriesDemoTest\ArmpsData\comEvaResult.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();

        }
        /// <summary>
        /// 把求得的W,B,w,b
        /// </summary>
        public void WB_Excel()
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();


            Workbooks workbooks = excelApp.Workbooks;
            Workbook workBook = workbooks.Add(Type.Missing);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1
            workSheet.Name = "w1";
            for (int i = 1; i <= this.w1.m; i++)
            {
                for (int j = 1; j <= this.w1.n; j++)
                {
                    workSheet.Cells[i, j] = this.w1.read(i - 1, j - 1);
                }
            }
            workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet = (Worksheet)workBook.Worksheets[1];
            workSheet.Name = "b1";
            for (int i = 1; i <= this.b1.m; i++)
            {
                for (int j = 1; j <= b1.n; j++)
                {
                    workSheet.Cells[i, j] = this.b1.read(i - 1, j - 1);
                }
            }

            workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet = (Worksheet)workBook.Worksheets[1];
            workSheet.Name = "w2";

            for (int i = 1; i <= this.w2.m; i++)
            {
                for (int j = 1; j <= w2.n; j++)
                {
                    workSheet.Cells[i, j] = this.w2.read(i - 1, j - 1);
                }
            }

            workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet = (Worksheet)workBook.Worksheets[1];
            workSheet.Name = "b2";

            for (int i = 1; i <= this.b2.m; i++)
            {
                for (int j = 1; j <= b2.n; j++)
                {
                    workSheet.Cells[i, j] = this.b2.read(i - 1, j - 1);
                }
            }




            workBook.SaveAs(@"f:\BPSeriesDemoTest\ArmpsData\saveWB.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();









        }
        /// <summary>
        /// 输出结果写入Excel
        /// </summary>
        public void output_Excel()
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();


            Workbooks workbooks = excelApp.Workbooks;
            Workbook workBook = workbooks.Add(Type.Missing);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1

            for (int j = 1; j < 5; j++)
            //workSheet.Cells[1, j] = accuracy[j - 1];
            {
                workSheet.Cells[j, 1] = output_test_Norm.read(j - 1, 0);
                workSheet.Cells[j, 2] = output_test_Norm.read(j - 1, 1);
                workSheet.Cells[j, 3] = output_test_Norm.read(j - 1, 2);
                workSheet.Cells[j, 4] = output_test_Norm.read(j - 1, 3);
                workSheet.Cells[j, 5] = fore_test.read(j - 1, 0);
                workSheet.Cells[j, 6] = fore_test.read(j - 1, 1);
                workSheet.Cells[j, 7] = fore_test.read(j - 1, 2);
                workSheet.Cells[j, 8] = fore_test.read(j - 1, 3);
                //workSheet.Cells[9, j] = accuracy[j - 1];
            }

            //workSheet.Cells[10, 1] = accu_average;

            workBook.SaveAs(@"f:\BPSeriesDemoTest\ArmpsData\result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();



        }
        /// <summary>
        /// 反归一化获得输出结果
        /// </summary>
        public void ConvNorm()
        {
            fore = matrix_Calc.transposs(this.fore_test);
            _Matrix output = matrix_Calc.transposs(this.Output_train);
            for (int i = 0; i < output.n; i++)
            {
                double max = -100000.0; double min = 100000.0;
                for (int j = 0; j < output.m; j++)
                {
                    if (max < output.read(j, i))
                    {
                        max = output.read(j, i);
                    }
                    else if (min > output.read(j, i))
                    {
                        min = output.read(j, i);
                    }

                }
                for (int j = 0; j < fore.m; j++)
                {
                    double s = (max - min) * fore.read(j, i) + min;
                    fore.write(j, i, s);
                }
            }
        }

        public void CalcAccuracy()
        {
            accuracy = new double[outnum];
            _Matrix output = matrix_Calc.transposs(output_train);
            accu_average = 0.0;
            for (int i = 0; i < outnum; i++)
            {
                double accu = 0.0;
                for (int j = 0; j < test_num; j++)
                {
                    accu += Math.Abs(fore.read(j, i) - output.read(j, i)) / output.read(j, i);
                    accu_average += Math.Abs(fore.read(j, i) - output.read(j, i)) / output.read(j, i);
                }
                accuracy[i] = accu / 10;
                accu_average = accu_average / (outnum) / (test_num);
            }

        }

    }
}
