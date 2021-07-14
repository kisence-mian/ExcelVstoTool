using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelVstoTool
{
    public partial class Ribbon_Main
    {
        private void Ribbon_Main_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button_initData_Click(object sender, RibbonControlEventArgs e)
        {
            //判断 config 页是否存在
            Worksheet config; 

            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, Const.c_SheetName_Config))
            {
                //创建一个Config 页面
                config = ExcelTool.CreateSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);

                DataConfig.ConfigInit(config);
            }
            else
            {
                config = Globals.ThisAddIn.Application.Worksheets[Const.c_SheetName_Config];
                MessageBox.Show(Const.c_SheetName_Config + "页面已经存在");
            }
        }

        #region 导入导出

        private void button_toTxt_Click(object sender, RibbonControlEventArgs e)
        {
            //先进行一次保存
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            Worksheet config = ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);

            //没有初始化直接返回
            if (config == null)
            {
                return;
            }

            //获取真实路径

            if(checkBox_exportCheck.Checked)
            {
                if(!CheckAll())
                {
                    MessageBox.Show("导出失败： 校验未通过");
                    return;
                }
            }

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.m_txtPath))
                {
                    Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);

                    DataTool.Excel2Data(wst, dataConfig);
                }
            }

            MessageBox.Show("导出完毕");
        }

        private void button_ToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            string info = "导入完毕";
            List<string> nofindPath = new List<string>();

            Worksheet config = ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);

            //没有初始化直接返回
            if(config == null)
            {
                return;
            }

            DateTime now = System.DateTime.Now;

            PerformanceSwitch(true);

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.m_txtPath))
                {
                    if(File.Exists(dataConfig.m_txtPath))
                    {
                        Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);
                        DataTool.Data2Excel(dataConfig, wst);
                    }
                    else
                    {
                        nofindPath.Add(dataConfig.m_txtPath);
                    }
                }
            }

            PerformanceSwitch(false);

            //构造输出信息
            info += "\n用时：" + (DateTime.Now - now).TotalSeconds + "s";

            //错误的路径配置
            if (nofindPath.Count > 0)
            {
                info += "\n找不到的文本";
                for (int i = 0; i < nofindPath.Count; i++)
                {
                    info += "\n  " + nofindPath[i];
                }
            }



            MessageBox.Show(info);
        }

        #endregion

        #region 数据

        private void button_check_Click(object sender, RibbonControlEventArgs e)
        {
            string info = "校验完毕";
            //先进行一次保存
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            Worksheet config = ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);

            //没有初始化直接返回
            if (config == null)
            {
                return;
            }

            DateTime now = System.DateTime.Now;

            //进行校验
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.m_txtPath))
                {
                    Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);
                    if(!CheckTool.CheckSheet(wst, dataConfig))
                    {
                        info += "\n->" + dataConfig.m_sheetName + " 校验未能通过";
                    }
                }
            }

            //构造输出信息
            info += "\n用时：" + (DateTime.Now - now).TotalSeconds + "s";

            MessageBox.Show(info);
        }

        bool CheckAll()
        {
            bool result = true;
            Worksheet config = ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);
            //进行校验
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.m_txtPath))
                {
                    Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);
                    
                    result &= CheckTool.CheckSheet(wst, dataConfig);
                }
            }
            return result;
        }

        private void button_dataInfo_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("初始化后可以使用对外部表格和多语言Key的校验");
        }

        #endregion

        #region 多语言

        private void button_dataInit_Click(object sender, RibbonControlEventArgs e)
        {
            DataManager.Init();
            LanguageInit();
        }

        void LanguageInit()
        {
            LanguageManager.Init();

            comboBox_currentLanguage.Enabled = LanguageManager.IsEnable;
            //gallery_language.Enabled = LanguageManager.IsEnable;
            button_changeLanguageColumn.Enabled = LanguageManager.IsEnable;

            box_dataInit.Visible = !LanguageManager.IsEnable;

            if(LanguageManager.IsEnable)
            {
                //下拉框
                comboBox_currentLanguage.Items.Clear();
                for (int i = 0; i < LanguageManager.allLanuage.Count; i++)
                {
                    RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                    tmp.Label = LanguageManager.allLanuage[i].ToString();
                    comboBox_currentLanguage.Items.Add(tmp);
                }

                comboBox_currentLanguage.Text = LanguageManager.currentLanguage.ToString();
            }
        }

        private void button_changeLanguageColumn_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("功能暂未实现");
        }

        private void button_LanguageInfo_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("此功能需要读取 Resources\\Data\\Language\\LanguageConfig.txt 的配置 \n 如果该配置不存在，功能不可用");
        }

        #endregion

        #region 工具方法

        //性能开关
        void PerformanceSwitch(bool enable)
        {
            if (checkBox_CloseView.Checked)
            {
                Globals.ThisAddIn.Application.Visible = !enable;
                Globals.ThisAddIn.Application.ScreenUpdating = !enable;
            }
        }



        #endregion


    }
}
