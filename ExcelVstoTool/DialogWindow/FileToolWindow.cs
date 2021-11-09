using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelVstoTool.DialogWindow
{
    public partial class FileToolWindow : Form
    {
        public FileToolWindow()
        {
            InitializeComponent();
        }

        private void button_batchExcel2Text_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new
            Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;//等于true表示可以选择多个文件
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel文件|*.xlsx";

            DateTime now = DateTime.Now;

            List<string> failList = new List<string>();

            string error = "";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in dlg.FileNames)
                {
                    string newName = FileTool.RemoveExpandName(file) + ".txt";

                    try
                    {
                        SaveAsToText(excelApp, file, newName);
                    }
                    catch(Exception ex)
                    {
                        failList.Add(file + " -> " + ex.ToString());
                    }
                }
            }

            excelApp.Quit();
            if(failList.Count > 0)
            {
                error += "转换错误的文件：\n";
                for (int i = 0; i < failList.Count; i++)
                {
                    error += failList[i] + "\n";
                }
            }

            MessageBox.Show("导出完毕\n用时：" + (DateTime.Now - now).TotalSeconds + "s\n" + error);
        }

        private void button_batchText2Excel_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;//等于true表示可以选择多个文件
            dlg.DefaultExt = ".txt";
            dlg.Filter = "Text文件|*.txt";

            DateTime now = DateTime.Now;
            List<string> failList = new List<string>();
            string error = "";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in dlg.FileNames)
                {
                    try
                    {
                        Text2Excel(Globals.ThisAddIn.Application, file);
                    }
                    catch(Exception ex)
                    {
                        failList.Add(file + " -> " + ex.ToString());
                    }
                }
            }

            if (failList.Count > 0)
            {
                error += "导入错误的文件：\n";
                for (int i = 0; i < failList.Count; i++)
                {
                    error += failList[i] + "\n";
                }
            }

            MessageBox.Show("导入完毕\n用时：" + (DateTime.Now - now).TotalSeconds + "s\n" + error);
        }

        void SaveAsToText(Microsoft.Office.Interop.Excel.Application excelApp, string excelPath,string txtPath)
        {
            Workbook workbook = null;

            workbook = excelApp.Workbooks.Open(excelPath);
            workbook.SaveAs(txtPath, XlFileFormat.xlUnicodeText);
            workbook.Close();
            excelApp.Quit();
        }

        void Text2Excel(Microsoft.Office.Interop.Excel.Application excelApp,string txtPath)
        {
            string sheetName = FileTool.RemoveExpandName( FileTool.GetFileNameBySring(txtPath));
            string content = FileTool.ReadStringByFile(txtPath);

            if(ExcelTool.ExistSheetName(excelApp, sheetName))
            {
                throw new Exception("已存在的工作簿名称 >" + sheetName + "<");
            }

            Worksheet sheet = ExcelTool.CreateSheet(excelApp, sheetName);

            List<List<string>> contentList = new List<List<string>>();

            string[] lines = content.Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                string[] values = line.Split('\t');
                for (int j = 0; j < values.Length; j++)
                {
                    sheet.Cells[1 + i, 1 + j] = values[j];
                }
            }
        }
    }
}
