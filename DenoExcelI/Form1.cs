using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace DenoExcelI
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Archivos Excel 97-2003 (.xls)|*.xls|Archivos Excel 2007 - (.xlsx)|*.xlsx";
            dlg.FilterIndex = 2;
            var result = dlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                string filename = dlg.FileName;
                try
                {
                    this.Create(@"c:\tmp\closexml.xlsx", filename);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                //dgRes.ItemsSource = helper.LoadDatafromExcel(filename).DefaultView;

                //helper.ImportFromExcel(filename);
                //helper.PerformBulkCopy(filename);
                //MessageBox.Show(this, string.Format("Importacion de datos completa "), "Sistema pronabes", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public void Create(String filePath, string OriFile)
        {
            var excelBook = new XLWorkbook(OriFile);
            var hoja = excelBook.Worksheet("2011-2012 Guadalupe A , Clement");
            var dataRange = hoja.RangeUsed();
            var dataTable = dataRange.AsTable();

            StringBuilder sb = new StringBuilder();
            StringBuilder sbCurp = new StringBuilder();
            List<string> lsCurps = new List<string>();


            foreach (var row in dataTable.Rows())
            {
                sb.Append(string.Format("{0},{1},{2},{3}",
                    row.Cell(1).GetString(), row.Cell(2).GetString(), row.Cell(3).GetString(), row.Cell(4).GetString()));
                sbCurp.AppendFormat("{0}", row.Cell("B").GetString());
                sb.Append("\n");
                sbCurp.Append("\n");
                lsCurps.Add(string.Format("{0}", row.Cell("B").GetString()));

                #region Updates
                
                int rNum = row.RowNumber();
                hoja.Cell(rNum, 7).Value = row.Cell("B").GetString();
                #endregion
            }

            using (StreamWriter outfile = new StreamWriter(@"C:\tmp\res.txt"))
            {
                outfile.Write(sb.ToString());
                outfile.Write(sbCurp.ToString());
            }

            updateCeldas(lsCurps);

            //MessageBox.Show(sb.ToString());

            //IXLWorksheet Hoja;
            //excelBook.TryGetWorksheet("2011-2012 Guadalupe A , Clement", out Hoja);
            //var wsActiveCell = Hoja;


            //var wsSelectRowsColumns = excelBook.AddWorksheet("Select Rows and Columns");
            //wsSelectRowsColumns.Rows("2, 4-15").Select();
            //wsSelectRowsColumns.Columns("2, 4-5").Select();

            //var wsSelectMisc = excelBook.AddWorksheet("Select Misc");
            //wsSelectMisc.Cell("B2").Select();
            //wsSelectMisc.Range("D2:E2").Select();
            //wsSelectMisc.Ranges("C3, D4:E5").Select();

            excelBook.SaveAs(filePath);
        }

        private void updateCeldas(List<string> lsCurps)
        {
            foreach (var row in lsCurps)
            {

            }
        }
    }
}
