using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
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
                window.SmallScroll(-worksheet.Rows.Count, Type.Missing, -worksheet.Columns.Count);
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
                window.SmallScroll(-worksheet.Rows.Count, Type.Missing, -worksheet.Columns.Count);
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
                fbd.Description = "対象フォルダ";
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath, "*.xls", SearchOption.AllDirectories);

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
                                app.ActiveWindow.SmallScroll(-worksheet.Rows.Count, Type.Missing, -worksheet.Columns.Count);
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
                    temp.get_Range("A" + i).Value = "処理完了　完了" + (i - 1);
                    MessageBox.Show("処理完了　完了" + (i - 1));
                }
            }

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Window window = e.Control.Context;
            Excel.Worksheet temp = (Excel.Worksheet)window.Application.ActiveSheet;
            int i = 1;
            var fbd = new FolderBrowserDialog();
            fbd.Description = "保存場所";
            if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                string path = fbd.SelectedPath;
                while (!string.IsNullOrEmpty(temp.Cells[i, 1].Value))
                {
                    string filename = temp.Cells[i, 1].Value;
                    filename = filename.Split('\\').Last();
                    filename = filename.Split('.').First();

                    Excel.Application app = new Excel.Application();
                    Excel.Workbook newworkbook = app.Workbooks.Add();
                    Excel.Worksheet newTab = newworkbook.Sheets[1];
                    newTab.Activate();
                    newTab.Cells[1, 1] = "Hello World";
                    newworkbook.SaveAs(path + "\\" + filename);
                    newworkbook.Close();
                    temp.Activate();
                    i++;
                }

                Process.Start(path);

            }
        }
    }
}
