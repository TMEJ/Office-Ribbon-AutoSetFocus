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
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
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
                try
                {
                    Excel.Worksheet worksheet = (Excel.Worksheet)item;
                    worksheet.Activate();
                    worksheet.get_Range("A1").Select();
                    window.SmallScroll(-worksheet.Rows.Count, Type.Missing, -worksheet.Columns.Count);
                }
                catch (Exception)
                {
                    continue;
                }
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
                try
                {
                    Excel.Worksheet worksheet = (Excel.Worksheet)item;
                    worksheet.Activate();
                    worksheet.get_Range("A1").Select();
                    window.SmallScroll(-worksheet.Rows.Count, Type.Missing, -worksheet.Columns.Count);
                }
                catch (Exception)
                {
                    continue;
                }
            }
            Excel.Worksheet sheet1 = (Excel.Worksheet)activeWorkbook.Sheets[1];
            sheet1.Activate();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Window window = e.Control.Context;

                Excel.Worksheet temp = new Excel.Worksheet();
                if (window == null)
                {
                    
                    Excel.Application newWindow = new Excel.Application();
                    newWindow.ShowWindowsInTaskbar = true;
                    newWindow.SheetsInNewWorkbook = 1;
                    Excel.Workbook newBook = newWindow.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    newBook.Activate();
                    temp = newBook.Sheets[1];
                    temp.Activate();
                }
                else
                {
                    temp = (Excel.Worksheet)window.Application.ActiveSheet;
                }

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
                                string messege = string.Empty;
                                try
                                {
                                    Excel.Application app = new Excel.Application();
                                    Excel.Workbook activeWorkbook = app.Workbooks.Open(item);
                                    foreach (var sheet in activeWorkbook.Sheets)
                                    {
                                        Excel.Worksheet worksheet = (Excel.Worksheet)sheet;
                                        worksheet.Activate();
                                        try
                                        {
                                            worksheet.get_Range("A1").Select();
                                        }
                                        catch (COMException ex)
                                        {
                                            if (ex.ErrorCode == -2146827284)
                                            {
                                                messege = "Has hidden pages or hidden rows.";
                                                continue;
                                            }
                                            else
                                            {
                                                throw;
                                            }
                                        }
                                        app.ActiveWindow.FreezePanes = false;
                                        app.ActiveWindow.SmallScroll(-worksheet.Rows.Count, Type.Missing, -worksheet.Columns.Count);
                                    }
                                    Excel.Worksheet sheet1 = (Excel.Worksheet)activeWorkbook.Sheets[1];
                                    sheet1.Activate();
                                    activeWorkbook.Save();
                                    activeWorkbook.Close();
                                    temp.Activate();
                                    temp.get_Range("A" + i).Value = item;
                                    if (messege != string.Empty)
                                    {
                                        temp.get_Range("B" + i).Value =  messege;
                                        temp.get_Range("B" + i).Font.Color = XlRgbColor.rgbOrange;
                                    }
                                    i++;

                                }
                                catch (COMException come)
                                {
                                    temp.Activate();
                                    temp.get_Range("A" + i).Value = item;
                                    temp.get_Range("B" + i).Value = come.Message + messege;
                                    temp.get_Range("B" + i).Font.Color = XlRgbColor.rgbRed;
                                    i++;
                                    continue;
                                }
                            }
                        }
                        temp.Activate();
                        temp.get_Range("A" + i).Value = "処理完了　完了" + (i - 1);
                        MessageBox.Show("処理完了　完了" + (i - 1));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
        }

        //private void button3_Click(object sender, RibbonControlEventArgs e)
        //{
        //    Excel.Window window = e.Control.Context;
        //    Excel.Worksheet temp = (Excel.Worksheet)window.Application.ActiveSheet;
        //    int i = 1;
        //    var fbd = new FolderBrowserDialog();
        //    fbd.Description = "保存場所";
        //    if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
        //    {
        //        string path = fbd.SelectedPath;
        //        while (!string.IsNullOrEmpty(temp.Cells[i, 1].Value))
        //        {
        //            string filename = temp.Cells[i, 1].Value;
        //            filename = filename.Split('\\').Last();
        //            filename = filename.Split('.').First();

        //            Excel.Application app = new Excel.Application();
        //            Excel.Workbook newworkbook = app.Workbooks.Add();
        //            Excel.Worksheet newTab = newworkbook.Sheets[1];
        //            newTab.Activate();
        //            newTab.Cells[1, 1] = "Hello World";
        //            newworkbook.SaveAs(path + "\\" + filename);
        //            newworkbook.Close();
        //            temp.Activate();
        //            i++;
        //        }

        //        Process.Start(path);

        //    }
        //}
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Window window = e.Control.Context;

                Excel.Worksheet temp = new Excel.Worksheet();
                if (window == null)
                {

                    Excel.Application newWindow = new Excel.Application();
                    newWindow.ShowWindowsInTaskbar = true;
                    newWindow.SheetsInNewWorkbook = 1;
                    Excel.Workbook newBook = newWindow.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    newBook.Activate();
                    temp = newBook.Sheets[1];
                    temp.Activate();
                }
                else
                {
                    temp = (Excel.Worksheet)window.Application.ActiveSheet;
                }

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

                                activeWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, item.Split('.')[0] + ".pdf", XlFixedFormatQuality.xlQualityStandard, false, false, Type.Missing, Type.Missing, false);
                                activeWorkbook.Close(false);
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //throw;
            }
        }
    }
}
