using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelVstoTool
{
    public partial class Ribbon_Main
    {
        #region 常量

        public string c_ConfigSheetName = "Config";

        #endregion

        string m_dataPath = "";

        private void Ribbon_Main_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button_initData_Click(object sender, RibbonControlEventArgs e)
        {
            //判断 config 页是否存在
            Worksheet config; 

            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, c_ConfigSheetName))
            {
                //创建一个Config 页面
                config = ExcelTool.CreateSheet(Globals.ThisAddIn.Application, c_ConfigSheetName);
                config.Range["A1"].Value = "页面名称";
                config.Range["B1"].Value = "导出表名";
            }
            else
            {
                config = Globals.ThisAddIn.Application.Worksheets[c_ConfigSheetName];
                MessageBox.Show(c_ConfigSheetName + "页面已经存在");
            }
        }

        private void button_toTxt_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet config = ExcelTool.GetSheet(Globals.ThisAddIn.Application, c_ConfigSheetName);

            //没有初始化直接返回
            if (config == null)
            {
                return;
            }

            //获取真实路径

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                string key = "A" + i;
                string value = "B" + i;

                string sheetName = config.Range[key].Value;
                string txtPath = config.Range[value].Value;

                if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(txtPath))
                {
                    Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, sheetName, true);

                    DataTool.Excel2Data(wst, txtPath);
                }
            }

            MessageBox.Show("导出完毕");
        }

        private void button_ToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet config = ExcelTool.GetSheet(Globals.ThisAddIn.Application, c_ConfigSheetName);

            //没有初始化直接返回
            if(config == null)
            {
                return;
            }

            //获取真实路径

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                string key = "A" + i;
                string value = "B" + i;

                string sheetName = config.Range[key].Value;
                string txtPath = config.Range[value].Value;

                if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(txtPath))
                {
                    Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, sheetName,true);

                    DataTool.Data2Excel(txtPath, wst);
                }
            }

            MessageBox.Show("导入完毕");
        }
    }
}
