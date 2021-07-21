using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Matrix_Mul;
using System.IO;
using NPOI.HSSF.UserModel;
using System.Data.OleDb;
using BPNETSerial;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace BP_Cal_dll
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public int InputLayerNum = 10;//20200414增加
        public int HiddenLayerNum = 15;//20200414增加
        public int OutputLayerNum = 10;//20200414增加

        public DataSet getData()//用于打开文件，打开已经过处理的数据文件，用于进行训练和预测，主要用于获取矩阵维度m和n
        {
            //打开文件
            OpenFileDialog file = new OpenFileDialog();
            //file.Filter = "Excel(*.xlsx)|*.xlsx|Excel(*.xls)|*.xls";
            file.Filter = "Excel(*.xls)|*.xls";
            file.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            file.Multiselect = false;
            if (file.ShowDialog() == DialogResult.Cancel)
                return null;
            //判断文件后缀
            var path = file.FileName;
            string fileSuffix = System.IO.Path.GetExtension(path);
            if (string.IsNullOrEmpty(fileSuffix))
                return null;

            using (DataSet ds = new DataSet())
            {
                //判断Excel文件是2003版本还是2007版本
                string connString = "";
                if (fileSuffix == ".xls")
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
                else
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + path + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
                //读取文件
                string sql_select = " SELECT * FROM [Sheet1$]";
                using (OleDbConnection conn = new OleDbConnection(connString))
                using (OleDbDataAdapter cmd = new OleDbDataAdapter(sql_select, conn))
                {
                    conn.Open();
                    cmd.Fill(ds);
                }
                if (ds == null || ds.Tables.Count <= 0) return null;
                return ds;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            DataTable dat = getData().Tables[0];//获取表中的所有数据，其中包含了电流特征数据、电流工况以及工况矩阵
            _Matrix Matrix_Feature = new _Matrix(dat.Rows.Count, InputLayerNum);//所有行，以及前10列为电流的特征数据//20200414从数值调整为变量，方便后续扩充
            _Matrix Matrix_Output = new _Matrix(dat.Rows.Count, dat.Columns.Count - InputLayerNum - 2);//所有行，除了前11列以外的数据为电流的输出矩阵//20200414从数值调整为变量，方便后续扩充
            Matrix_Feature.init_matrix();
            Matrix_Output.init_matrix();
            OutputLayerNum = Matrix_Output.n;//20210716扩充工况的bug修复
            using (FileStream stream = new FileStream(@"TrainMatrix-13工况.xls", FileMode.Open, FileAccess.Read))//打开训练集数据
            {
                HSSFWorkbook Workbook = new HSSFWorkbook(stream);
                var Sheet = Workbook.GetSheetAt(0);
                int j = 0;
                for (int i = 1; i <= dat.Rows.Count; i++)
                {
                    var row = Sheet.GetRow(i);
                    for (int k = 0; k < InputLayerNum; k++)//20200414从数值调整为变量，方便后续扩充
                    {
                        Matrix_Feature.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }

                j = 0;
                for (int i = 1; i <= dat.Rows.Count; i++)
                {
                    var row = Sheet.GetRow(i);
                    for (int k = InputLayerNum + 2; k < dat.Columns.Count; k++)//0的工况不作为训练的种类之一//20200414从数值调整为变量，方便后续扩充
                    {
                        Matrix_Output.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }


        }//导入训练数据进行训练




            //初始化bp神经网络，参数分别为输入层节点数，隐层节点，输出层节点，数据样本数，数据维度（对应输入层节点数），迭代步数，学习率
            BP bp = new BP(InputLayerNum, HiddenLayerNum, OutputLayerNum, dat.Rows.Count, OutputLayerNum, 2000, 0.2);//此处有修改  可结合输出层节点数确定网络结构//20200414从数值调整为变量，方便后续扩充
            //创建一个mat来便于对_Matrix类进行计算
            _Matrix_Calc mat = new _Matrix_Calc();

            bp.Input_train = mat.transposs(Matrix_Feature);
            _Matrix Output_train = new _Matrix(Matrix_Output);
            bp.Output_train = mat.transposs(Matrix_Output);
            bp.trainBP(bp.Input_train, bp.Output_train);

            WB_Excel(bp.w1, bp.w2, bp.b1, bp.b2);//保存训练的权重矩阵



            //output_Excel(bp.output_test_Norm, bp.fore_test);//output的此种重载用于输出评价矩阵和评价结果两个内容，与上一个只输出结果不同，调试用
        }

        private void button2_Click(object sender, EventArgs e)
        {

            DataTable dat2 = getData().Tables[0];//获取表中的所有数据，其中包含了电流特征数据、电流工况以及工况矩阵
            _Matrix Matrix_Feature_Pred = new _Matrix(dat2.Rows.Count, InputLayerNum);//所有行，以及前10列为电流的特征数据//20200414从数值调整为变量，方便后续扩充
            _Matrix Matrix_Output_Pred = new _Matrix(dat2.Rows.Count, dat2.Columns.Count - InputLayerNum - 2);//所有行，除了前11列以外的数据为电流的输出矩阵//20200414从数值调整为变量，方便后续扩充
            Matrix_Feature_Pred.init_matrix();
            Matrix_Output_Pred.init_matrix();
            OutputLayerNum = Matrix_Output_Pred.n;//20210716扩充工况类型的变量修复



            using (FileStream stream = new FileStream(@"TrainMatrix-13工况.xls", FileMode.Open, FileAccess.Read))//打开预测集数据
            {
                HSSFWorkbook Workbook = new HSSFWorkbook(stream);
                var Sheet = Workbook.GetSheetAt(0);
                int j = 0;
                for (int i = 1; i <= dat2.Rows.Count; i++)
                {
                    var row = Sheet.GetRow(i);
                    for (int k = 0; k < InputLayerNum; k++)//20200414从数值调整为变量，方便后续扩充//20210716修正输入输出的bug数值
                    {
                        Matrix_Feature_Pred.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }
                j = 0;
                for (int i = 1; i <= dat2.Rows.Count; i++)
                {
                    var row = Sheet.GetRow(i);
                    for (int k = InputLayerNum + 2; k < dat2.Columns.Count; k++)//0的工况不作为训练的种类之一//20200414从数值调整为变量，方便后续扩充
                    {
                        Matrix_Output_Pred.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }
            }//导入预测数据进行预测

             //初始化bp神经网络，参数分别为输入层节点数，隐层节点，输出层节点，数据样本数，数据维度（对应输入层节点数），迭代步数，学习率
            BP bp = new BP(InputLayerNum, HiddenLayerNum, OutputLayerNum, dat2.Rows.Count, OutputLayerNum, 2000, 0.2);//20200414从数值调整为变量，方便后续扩充
            //创建一个mat来便于对_Matrix类进行计算
            _Matrix_Calc mat = new _Matrix_Calc();
            bp.b2.init_matrix();
            bp.b1.init_matrix();
            bp.w1.init_matrix();
            bp.w2.init_matrix();

            using (FileStream stream = new FileStream(@"saveWB.xls", FileMode.Open, FileAccess.Read))//打开权重矩阵
            {
                HSSFWorkbook Workbook = new HSSFWorkbook(stream);
                var Sheet = Workbook.GetSheetAt(0);
                int j = 0;
                for (int i = 0; i < OutputLayerNum; i++)//20200414从数值调整为变量，方便后续扩充
                {
                    var row = Sheet.GetRow(i);
                    for (int k = 0; k < 1; k++)
                    {
                        double a = row.GetCell(k).NumericCellValue;
                        bp.b2.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }
                var Sheet2 = Workbook.GetSheetAt(1);
                j = 0;
                for (int i = 0; i < HiddenLayerNum; i++)//20200414从数值调整为变量，方便后续扩充
                {
                    var row = Sheet2.GetRow(i);
                    for (int k = 0; k < OutputLayerNum; k++)//20200414从数值调整为变量，方便后续扩充
                    {
                        bp.w2.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }
                var Sheet3 = Workbook.GetSheetAt(2);
                j = 0;
                for (int i = 0; i < HiddenLayerNum; i++)//20200414从数值调整为变量，方便后续扩充
                {
                    var row = Sheet3.GetRow(i);
                    for (int k = 0; k < 1; k++)
                    {
                        bp.b1.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }
                var Sheet4 = Workbook.GetSheetAt(3);
                j = 0;
                for (int i = 0; i < HiddenLayerNum; i++)//20200414从数值调整为变量，方便后续扩充
                {
                    var row = Sheet4.GetRow(i);
                    for (int k = 0; k < InputLayerNum; k++) //20200414从数值调整为变量，方便后续扩充
                    {
                        bp.w1.arr[j++] = row.GetCell(k).NumericCellValue;
                    }
                }

            }//导入权重数据进行计算





            bp.Input_test = mat.transposs(Matrix_Feature_Pred);
            bp.Output_test = mat.transposs(Matrix_Output_Pred);
            bp.testBP(bp.Input_test, bp.Output_test, dat2.Rows.Count);

            bp.ConvNorm();
            bp.CalcAccuracy();
            //bp.GetComprehensiveEvaluationMatrix(10);
            bp.GetComprehensiveEvaluationMatrix(4);
            bp.comprehensiveEvaluation = new double[4] { 0.3009, 0.2201, 0.2696, 0.1793 };
            bp.ComEvaluation(bp.comprehensiveEvaluation);

            //bp.ComEvaResult_Excel();
            //ComEvaResult_Excel(bp.comprehensiveEvaluationResultMatrix);调试用，实际使用中不用输出
            //bp.WB_Excel();
            //WB_Excel(bp.w1, bp.w2, bp.b1, bp.b2);//调试用，实际使用中不用输出
            //bp.output_Excel();

            //_Matrix Result_Matrix = Get_Result(bp.output_test_Norm, bp.fore_test);//20200414隐去

            //output_Excel(Result_Matrix);//20200414隐去
            output_Excel(bp.output_test_Norm, bp.fore_test);//output的此种重载用于输出评价矩阵和评价结果两个内容，与上一个只输出结果不同，调试用


        }

        /// <summary>
        /// ComEvaResult矩阵写入EXCEL////调试用，正常使用时可隐去本函数
        /// </summary>
        public void ComEvaResult_Excel(_Matrix comprehensiveEvaluationResultMatrix)
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

            workBook.SaveAs(@"c:\BPSeriesDemoTest\ArmpsData\comEvaResult.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();

        }


        /// <summary>
        /// 把求得的W,B,w,b写入EXCEL//调试用，正常使用时可隐去本函数
        /// </summary>
        public void WB_Excel(_Matrix w1, _Matrix w2, _Matrix b1, _Matrix b2)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();


            Workbooks workbooks = excelApp.Workbooks;
            Workbook workBook = workbooks.Add(Type.Missing);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1
            workSheet.Name = "w1";
            for (int i = 1; i <= w1.m; i++)
            {
                for (int j = 1; j <= w1.n; j++)
                {
                    workSheet.Cells[i, j] = w1.read(i - 1, j - 1);
                }
            }
            workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet = (Worksheet)workBook.Worksheets[1];
            workSheet.Name = "b1";
            for (int i = 1; i <= b1.m; i++)
            {
                for (int j = 1; j <= b1.n; j++)
                {
                    workSheet.Cells[i, j] = b1.read(i - 1, j - 1);
                }
            }

            workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet = (Worksheet)workBook.Worksheets[1];
            workSheet.Name = "w2";

            for (int i = 1; i <= w2.m; i++)
            {
                for (int j = 1; j <= w2.n; j++)
                {
                    workSheet.Cells[i, j] = w2.read(i - 1, j - 1);
                }
            }

            workBook.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet = (Worksheet)workBook.Worksheets[1];
            workSheet.Name = "b2";

            for (int i = 1; i <= b2.m; i++)
            {
                for (int j = 1; j <= b2.n; j++)
                {
                    workSheet.Cells[i, j] = b2.read(i - 1, j - 1);
                }
            }
            workBook.SaveAs(@"c:\BPSeriesDemoTest\ArmpsData\saveWB.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();
        }


        /// <summary>
        /// 输出结果写入Excel（此为连矩阵带结论，调试用，正常使用时可以隐去本函数）
        /// </summary>
        public void output_Excel(_Matrix output_test_Norm, _Matrix fore_test)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();


            Workbooks workbooks = excelApp.Workbooks;
            Workbook workBook = workbooks.Add(Type.Missing);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1
            fore_test = result_process(fore_test);//20200414增加
            for (int i = 1; i < fore_test.m + 1; i++)
            //workSheet.Cells[1, j] = accuracy[j - 1];
            {
                double max = -100;
                int[] flag = new int[fore_test.m];
                for (int j = 1; j < fore_test.n + 1; j++)
                {
                    workSheet.Cells[i, j] = fore_test.read(i - 1, j - 1);
                    if (fore_test.read(i - 1, j - 1) > max)
                    {
                        max = fore_test.read(i - 1, j - 1);
                        flag[i - 1] = j;
                    }
                }
                ////20210720补充用于检测启停泵当天的数据，设计逻辑：当输出数据一行中每列都相等则为该工况
                if (Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 1)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 2)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 3)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 4)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 5)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 6)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 7)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 8)) < 0.00001 &&
                    Math.Abs(fore_test.read(i - 1, 0) - fore_test.read(i - 1, 9)) < 0.00001)
                    flag[i - 1] = 11;
                /*
                for (int j = 1; j < fore_test.n + 1; j++)
                {
                    workSheet.Cells[i, j] = fore_test.read(i - 1, j - 1);
                    if (fore_test.read(i - 1, j - 1) > max)
                    {
                        max = fore_test.read(i - 1, j - 1);
                        flag[i - 1] = j;
                    }
                }
                */
                ////20210720补充部分结束
                workSheet.Cells[i, fore_test.n + 1] = flag[i - 1];
            }
            //workSheet.Cells[10, 1] = accu_average;
            workBook.SaveAs(@"c:\BPSeriesDemoTest\ArmpsData\result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();
        }

        /// <summary>
        /// 从输出矩阵中计算最后结果，与单参数传参重载的output_Excel函数配套使用，实现上边双参数传参重载的output_Excel函数实现相同的功能
        /// </summary>
        public _Matrix Get_Result(_Matrix output_test_Norm, _Matrix fore_test)
        {
            _Matrix Result = new _Matrix(fore_test.m, 1);
            Result.init_matrix();
            for (int i = 0; i < fore_test.m; i++)
            //workSheet.Cells[1, j] = accuracy[j - 1];
            {
                double max = -100;
                for (int j = 0; j < fore_test.n; j++)
                {
                    if (fore_test.read(i, j) > max)
                    {
                        max = fore_test.read(i, j);
                        Result.write(i, 0, j + 1);
                    }
                }
            }
            return Result;
        }

        /// <summary>
        /// 输出结果写入Excel（将最终结果输出到excel中）
        /// </summary>
        public void output_Excel(_Matrix result_Matrix)
        {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();

            Workbooks workbooks = excelApp.Workbooks;
            Workbook workBook = workbooks.Add(Type.Missing);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1
            result_Matrix = result_process(result_Matrix);//20200414增加
            for (int i = 1; i < result_Matrix.m + 1; i++)
            {
                workSheet.Cells[i, 1] = result_Matrix.read(i - 1, 0);
            }
            workBook.SaveAs(@"c:\BPSeriesDemoTest\ArmpsData\result.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();


        }

        /// <summary>
        /// 输出结果的尾部处理（将隶属度函数处理成概率）//20200414增加
        /// </summary>
        public _Matrix result_process(_Matrix result_Matrix)
        {
            for(int i = 0; i < result_Matrix.m; i++)
            {
                double sum = 0;
                for(int j = 0; j < result_Matrix.n; j++)
                {
                    if (result_Matrix.read(i, j) < 0)
                    {
                        result_Matrix.write(i, j, 0);
                    }
                    sum += result_Matrix.read(i, j);
                }
                for(int j = 0; j < result_Matrix.n; j++)
                {
                    result_Matrix.write(i, j, result_Matrix.read(i, j) / sum);
                }
            }

            return result_Matrix;
        }
    }
}
