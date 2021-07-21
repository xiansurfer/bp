using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Collections;
using ArmpsCardProcessing;
using Matrix_Mul;
using NPOI.HSSF.UserModel;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;





namespace ArmpsCard_dll
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonArmpsCardIO_Click(object sender, EventArgs e)
        {
            ClassArmpsCardProcess CACP = new ClassArmpsCardProcess();


            using (FileStream stream = new FileStream(@"ArmpsCardData12工况.xls", FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook Workbook = new HSSFWorkbook(stream);                
                var Sheet = Workbook.GetSheetAt(0);
                var row = Sheet.GetRow(1);
                int Columnofdata = row.Cells.Count;
                CACP.Input_A = new _Matrix( 241, Columnofdata);//20200402修改
                CACP.Input_A.init_matrix();
                CACP.Output_A = new _Matrix(1, Columnofdata);//20200402修改
                CACP.Output_A.init_matrix();
                //******Input中存储输入的电流数据*******//
                int j = 0;
                for (int i = 1; i < 242; i++)//导入excel的2-242行数据作为输入数据
                {
                    var row1 = Sheet.GetRow(i);//获取列总数，是每种电流卡片的样本数量
                    for (int k = 0; k < row1.Cells.Count; k++)
                    {
                        CACP.Input_A.arr[j++] = row1.GetCell(k).NumericCellValue;
                    }
                }
                //*****Output中存储输出结果****//
                j = 0;
                var row2 = Sheet.GetRow(0);//获取列总数，是每种电流卡片的样本数量
                for (int k = 0; k < row2.Cells.Count; k++)
                {
                    CACP.Output_A.arr[j++] = row2.GetCell(k).NumericCellValue;//存储电流卡片中的预设结论
                }
            }
            _Matrix MixtureMatrix = CACP.DataPreprocessingAndPrint(CACP.Input_A, CACP.Output_A);//处理数据形成特征矩阵，附带一列结果向量
            _Matrix OutputMatrix = CACP.OutputProcessing(CACP.Output_A);//处理完成后形成基于结果的数据矩阵
            output_Excel_DataProcessing(CACP.Input_A, MixtureMatrix, OutputMatrix, CACP.WorkingCondition);            
        }


        /// <summary>
        /// 处理结果输出写入Excel
        /// </summary>
        public void output_Excel_DataProcessing(_Matrix Input_A, _Matrix MixtureMatrix, _Matrix OutputMatrix,int WorkingCondition)            
        {                
            var excelApp = new Microsoft.Office.Interop.Excel.Application();               
            Workbooks workbooks = excelApp.Workbooks;                
            Workbook workBook = workbooks.Add(Type.Missing);              
            Worksheet workSheet = (Worksheet)workBook.Worksheets[1];//取得sheet1           
            for (int i = 1; i <= Input_A.n; i++)//Input_A只用于读取数据维度，以便确定所需输出的数据维度
            {
                for (int j = 1; j < 12; j++)//从0-9共10种输入参数，以及输出为第11种参数，共放入1-11列中，特征参数+结果参数
                    workSheet.Cells[i, j] = MixtureMatrix.read(i - 1, j - 1);
                for (int k = 1; k <= WorkingCondition; k++)//从0-10共11种工况 写入1-11列中
                    workSheet.Cells[i, k + 11] = OutputMatrix.read(k - 1, i - 1);//此处有修改20200402
            }            
            workBook.SaveAs(@"c:\BPSeriesDemoTest\ArmpsData\NormMatrix.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workbooks.Close();           
        }
    }
}

