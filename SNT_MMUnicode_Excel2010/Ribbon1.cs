using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Tools.Excel;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace SNT_MMUnicode_Excel2010
{
    public partial class Ribbon1
    {
        string output, input;
        object unknownType = Type.Missing;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ZGToUni_Click(object sender, RibbonControlEventArgs e)
        {
            //  Microsoft.Office.Interop.Excel.Application excel = Globals.ThisAddIn.Application;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range excelRange = Globals.ThisAddIn.Application.Selection;
            int RowCount = excelRange.Rows.Count;
            int ColumnCount = excelRange.Columns.Count;
            //
            //MessageBox.Show("Row count" + RowCount+"Column Count" + ColumnCount);
            UpdateCellValue(RowCount, ColumnCount, excelRange, "Pyidaungsu");
            /* input = excelRange.Value;
             output = Rabbit.Zg2Uni(input);
             excelRange.Value = output;
             excelRange.Font.Name = "Pyidaungsu";*/
        }
        private void UpdateCellValue(int RowCount, int ColumnCount, Excel.Range excelRange,string FontName)
        {



            for (int r = 1; r <= RowCount; r++)
            {
                for (int c = 1; c <= ColumnCount; c++)
                {
                    dynamic cell = excelRange.Cells[r, c];

                    try
                    {
                        string content = cell.Value2;
                        output = Rabbit.Zg2Uni(content);
                        cell.Value2 = output;

                        cell.Font.Name = FontName;

                        /* if (cell.Locked == false)
                         {
                             string content = cell.Value2;
                           if (content != null && !content.Trim().Equals(""))
                            {
                                content = content.Trim();
                                cell.Value2 = cell.Value2 + " - This is a test";
                            }
                          
                           
                            
                         } */
                    }
                    catch (Exception)
                    {
                        // we are using dynamic type for cell variable so
                        // the variable might not have all the properties we used in our code
                    }

                }
            }




        }

        private void btnSavePdf_Click(object sender, RibbonControlEventArgs e)
        {
            //string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string fileName = "QuickExport.pdf";
            // SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "PDF Files|*.pdf";
            saveFileDialog1.Title = "Save a PDF File";
            saveFileDialog1.InitialDirectory = @"D:\";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var file = new FileInfo(saveFileDialog1.FileName);
                Globals.ThisAddIn.Application.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                     Path.Combine(file.DirectoryName, saveFileDialog1.FileName), unknownType,true);
            }
        }

        private void btnXPS_Click(object sender, RibbonControlEventArgs e)
        {
            //string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string fileName = "QuickExport.pdf";
            // SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "PDF Files|*.pdf";
            saveFileDialog1.Title = "Save a PDF File";
            saveFileDialog1.InitialDirectory = @"D:\";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var file = new FileInfo(saveFileDialog1.FileName);
                Globals.ThisAddIn.Application.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypeXPS,
                     Path.Combine(file.DirectoryName, saveFileDialog1.FileName), unknownType,
                    true);
            }
        }

        private void U_Zaw_Click(object sender, RibbonControlEventArgs e)
        {
            //  Microsoft.Office.Interop.Excel.Application excel = Globals.ThisAddIn.Application;
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range excelRange = Globals.ThisAddIn.Application.Selection;
            int RowCount = excelRange.Rows.Count;
            int ColumnCount = excelRange.Columns.Count;
            //
            //MessageBox.Show("Row count" + RowCount+"Column Count" + ColumnCount);
            UpdateCellValue(RowCount, ColumnCount, excelRange, "Zawgyi-One");
            /* input = excelRange.Value;
             output = Rabbit.Zg2Uni(input);
             excelRange.Value = output;
             excelRange.Font.Name = "Pyidaungsu";*/
        }
    }
}
