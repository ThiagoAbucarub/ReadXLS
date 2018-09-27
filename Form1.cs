
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadXLS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dataTable = new DataTable();

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Thiago\Downloads\FIDC Venda de veículos Nissan 14092018 - Por originador.xlsx");//, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel._Worksheet xlWorksheet = (Excel._Worksheet)xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int j = 1; j <= colCount; j++)
            {
                string data = xlWorksheet.Cells[1, j].Value;
                dataTable.Columns.Add(data);
            }

            for (int i = 2; i <= rowCount; i++)
            {
                DataRow drRow = dataTable.NewRow();

                for (int j = 1; j <= colCount; j++)
                {
                    drRow[j - 1] = Convert.ToString(xlWorksheet.Cells[i, j].Value2);
                        
                    
                    
                }
                dataTable.Rows.Add(drRow);
            }

            xlWorkbook.Close();
            xlApp.Quit();

            dataGridView1.DataSource = dataTable;
        }
    }
}
