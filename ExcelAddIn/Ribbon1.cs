using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
namespace ExcelAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void autoFocusButton_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;
            
            Excel.Workbook activeWorkbook = (Excel.Workbook)window.Application.ActiveWorkbook;

            foreach (var item in activeWorkbook.Sheets)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)item;
                worksheet.Activate();
                worksheet.get_Range("A1").Select();
            }
            Excel.Worksheet sheet1 = (Excel.Worksheet)activeWorkbook.Sheets[1];
            sheet1.Activate();
            activeWorkbook.Save();
            try
            {
                if (window.Application.Workbooks.Count == 1)
                {
                    window.Close();
                }
                else
                {
                    activeWorkbook.Close();

                }
            }
            catch (Exception)
            {
                window.Close();
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;

            Excel.Workbook activeWorkbook = (Excel.Workbook)window.Application.ActiveWorkbook;

            foreach (var item in activeWorkbook.Sheets)
            {
                Excel.Worksheet worksheet = (Excel.Worksheet)item;
                worksheet.Activate();
                worksheet.get_Range("A1").Select();
            }
            Excel.Worksheet sheet1 = (Excel.Worksheet)activeWorkbook.Sheets[1];
            sheet1.Activate();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;
            Excel.Worksheet temp = (Excel.Worksheet)window.Application.ActiveSheet;
            
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath,"*.xls",SearchOption.AllDirectories);

                    int i = 1;
                    foreach (var item in files)
                    {
                        if (File.Exists(item))
                        {
                            Excel.Application app = new Excel.Application();
                            Excel.Workbook activeWorkbook = app.Workbooks.Open(item);
                            foreach (var sheet in activeWorkbook.Sheets)
                            {
                                Excel.Worksheet worksheet = (Excel.Worksheet)sheet;
                                worksheet.Activate();
                                worksheet.get_Range("A1").Select();
                            }
                            Excel.Worksheet sheet1 = (Excel.Worksheet)activeWorkbook.Sheets[1];
                            sheet1.Activate();
                            activeWorkbook.Save();
                            activeWorkbook.Close();
                            temp.Activate();
                            temp.get_Range("A" + i).Value = item;
                            i++;
                        }
                    }
                    temp.Activate();
                    temp.get_Range("A" + i).Value = "処理完了　完了" + i;

                }
            }

        }
    }
}
