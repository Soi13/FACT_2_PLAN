using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;


namespace Svod
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application ObjExcel;
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label1.Text = openFileDialog1.FileName;
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            /////Создание объекта Факт
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets["ФАКТ"];
            /////////

            /////Создание объекта План
            Microsoft.Office.Interop.Excel.Application ObjExcel1 = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook1 = ObjExcel1.Workbooks.Open(openFileDialog2.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet1;
            ObjWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook1.Sheets["ФАКТ"];
            /////////

            //Перенос строк с 17 по 20 в столбце №7
            ObjWorkSheet1.Cells[17, 7].value = ObjWorkSheet.Cells[17, 7];
            ObjWorkSheet1.Cells[18, 7].value = ObjWorkSheet.Cells[18, 7];
            ObjWorkSheet1.Cells[19, 7].value = ObjWorkSheet.Cells[19, 7];
            ObjWorkSheet1.Cells[20, 7].value = ObjWorkSheet.Cells[20, 7];
            
            
            //////////////////Заполнение ПЛАН В/////////////////////////////
            int nn = 1;
            progressBar1.Maximum = 338;
            label5.Text = "Перенос факта в план";

            //Цикл по БДДС подразделения  
            for (int s = 22; s<=338; s++)
            {
                /*double stolbec_summa_bdds_podr = Convert.ToDouble(ObjWorkSheet.Cells[s, 5].value);
                if (stolbec_summa_bdds_podr > 0)
                {
                    double stolbec_summa_svod = Convert.ToDouble(ObjWorkSheet1.Cells[s, 5].value);
                    double sum_common = stolbec_summa_svod + stolbec_summa_bdds_podr;
                    ObjWorkSheet1.Cells[s, 5] = sum_common;

                }*/
                
                for (int d = 10; d <= 41; d++)
                 {
                     //string date_bdds_podr = ObjWorkSheet.Cells[3, d].Text;
                     double summa_bdds = Convert.ToDouble(ObjWorkSheet.Cells[s, d].value);

                     if (summa_bdds>0)
                     {
                         double st = Convert.ToDouble(ObjWorkSheet1.Cells[s, d].value);
                         double sum = st + summa_bdds;
                         ObjWorkSheet1.Cells[s, d] = sum;
                     }
         
                 }
                nn = nn + 1;
                progressBar1.Value = nn;
                label4.Text ="Обработано строк: " + Convert.ToString(nn);
            }
            ///////////////////////////////////////////////////////////////////
            
            
            ObjWorkBook.Close();
            ObjExcel.Quit();
            ObjWorkBook = null;
            ObjWorkSheet = null;
            ObjExcel = null;

            ObjExcel1.Visible = true;
            
            GC.Collect(); 


        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {                
                    label3.Text = openFileDialog2.FileName;
            }
        }
    }
}
