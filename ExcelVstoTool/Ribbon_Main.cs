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
            //添加事件监听
            Globals.ThisAddIn.Application.WorkbookActivate += Application_WorkbookActivate; //打开工作簿
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate; //激活页签
            Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;//选中区域
        }

        private void Application_WorkbookActivate(Workbook Wb)
        {
            //如果读取到设置页签则自动进行初始化
            Worksheet config = GetConfigSheet();
            if (config != null)
            {
                //隐藏初始化按钮
                button_initConfig.Visible = false;

                if (!DataManager.IsEnable)
                {
                    DataInit();
                    UpdateDataUI();
                }

                if (!LanguageManager.IsEnable)
                {
                    LanguageInit();
                }
            }
        }

        private void Application_SheetActivate(object Sh)
        {
            Data_OnSheetChange(Globals.ThisAddIn.Application.ActiveSheet);
            Language_OnSheetChange(Globals.ThisAddIn.Application.ActiveSheet);
        }

        private void Application_SheetSelectionChange(object Sh, Range Target)
        {
            Data_OnSelectChange(Globals.ThisAddIn.Application.ActiveSheet, Target);
            Language_OnSelectChange(Globals.ThisAddIn.Application.ActiveSheet,Target);
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

                button_initConfig.Visible = false;
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
            Worksheet config = GetConfigSheet();

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

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
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

            Worksheet config = GetConfigSheet();

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

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
                {
                    if(File.Exists(dataConfig.GetTextPath()))
                    {
                        Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);
                        DataTool.Data2Excel(dataConfig, wst);
                    }
                    else
                    {
                        nofindPath.Add(dataConfig.GetTextPath());
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

        private void button_dataInit_Click(object sender, RibbonControlEventArgs e)
        {
            DataInit();
            LanguageInit();
        }

        void DataInit()
        {
            DataManager.Init();

            //构造类型的下拉列表
            dropDown_dataType.Items.Clear();
            foreach (FieldType fieldType in Enum.GetValues(typeof( FieldType)))
            {
                RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                tmp.Label = fieldType.ToString();
                dropDown_dataType.Items.Add(tmp);
            }

            //用途的下拉列表
            dropDown_assetsType.Items.Clear();

            foreach (DataFieldAssetType assetsType in Enum.GetValues(typeof(DataFieldAssetType)))
            {
                RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                tmp.Label = assetsType.ToString();
                dropDown_assetsType.Items.Add(tmp);
            }

            //根据类型或者用途判断次级类型是否显示
        }

        private void button_check_Click(object sender, RibbonControlEventArgs e)
        {
            string info = "校验完毕";
            //先进行一次保存
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            Worksheet config = GetConfigSheet();

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

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
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
            Worksheet config = GetConfigSheet();
            //进行校验
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
                {
                    Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);
                    
                    result &= CheckTool.CheckSheet(wst, dataConfig);
                }
            }
            return result;
        }

        private void button_generateDataClass_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button_dataInfo_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("初始化后可以使用对外部表格和多语言Key的校验");
        }

        private void dropDown_dataType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            DataManager.CurrentFieldType = (FieldType)Enum.Parse(typeof(FieldType), dropDown_dataType.SelectedItem.Label);
            ResetTypeString();
        }

        private void dropDown_assetsType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            DataManager.CurrentAssetType = (DataFieldAssetType)Enum.Parse(typeof(DataFieldAssetType), dropDown_assetsType.SelectedItem.Label);

            ResetTypeString();
        }

        void ResetSecTypeDropDownItem()
        {
            dropDown_secType.Items.Clear();

            if (DataManager.CurrentAssetType == DataFieldAssetType.TableKey)
            {
                List<string> list = DataManager.TableName;

                for (int i = 0; i < list.Count; i++)
                {
                    RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                    tmp.Label = list[i];
                    dropDown_secType.Items.Add(tmp);
                }
            }
            else if(DataManager.CurrentFieldType == FieldType.Enum)
            {
                List<string> list = DataManager.EnumName;

                for (int i = 0; i < list.Count; i++)
                {
                    RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                    tmp.Label = list[i];
                    dropDown_secType.Items.Add(tmp);
                }
            }
            else
            {
                RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                tmp.Label = "";
                dropDown_secType.Items.Add(tmp);
            }

        }

        private void dropDown_secType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            DataManager.CurrentSecType = dropDown_secType.SelectedItem.Label;
            ResetTypeString();
        }

        private void Data_OnSheetChange(Worksheet activeSheet)
        {
            if (DataManager.IsEnable)
            {
                //更新UI
                UpdateDataUI();
            }
        }

        private void Data_OnSelectChange(Worksheet sheet,Range target)
        {
            if (DataManager.IsEnable)
            {
                if (IsConfigWorkSheet())
                {
                    //第一列不处理
                    if(target.Column == 1)
                    {
                        DataManager.PaseToCurrentType(null);
                        UpdateDataUI();
                        return;
                    }


                    string keyString = sheet.Cells[1, target.Column].Text;
                    string typeString = sheet.Cells[2, target.Column].Text;

                    //先判断Key存不存在
                    if (!string.IsNullOrEmpty(keyString))
                    {
                        //再判断type存不存在
                        if (!string.IsNullOrEmpty(typeString))
                        {
                            DataManager.PaseToCurrentType(typeString);
                        }
                        else
                        {
                            //不存在当默认值处理
                            DataManager.PaseToCurrentType("");
                        }
                    }
                    else
                    {
                        DataManager.PaseToCurrentType(null);
                    }
                }
                else
                {
                    DataManager.PaseToCurrentType(null);
                }

                //更新UI
                UpdateDataUI();
            }
        }

        void UpdateDataUI()
        {
            button_generateDataClass.Enabled = IsConfigWorkSheet();

            ResetSecTypeDropDownItem();

            SetDropDown(dropDown_dataType, DataManager.CurrentFieldType.ToString());
            SetDropDown(dropDown_assetsType, DataManager.CurrentAssetType.ToString());
            SetDropDown(dropDown_secType, DataManager.CurrentSecType.ToString());

            if (!DataManager.isWorkRange)
            {
                dropDown_dataType.Enabled = false;
                dropDown_secType.Enabled = false;
                dropDown_assetsType.Enabled = false;

                return;
            }

            dropDown_dataType.Enabled = true;
            if (DataManager.CurrentFieldType == FieldType.String)
            {
                dropDown_assetsType.Enabled = true;
                
                if (DataManager.CurrentAssetType == DataFieldAssetType.TableKey)
                {
                    dropDown_secType.Enabled = true;
                }
                else
                {
                    dropDown_secType.Enabled = false;
                }
            }
            else if (DataManager.CurrentFieldType == FieldType.Enum)
            {
                dropDown_secType.Enabled = true;
                dropDown_assetsType.Enabled = false;
            }
            else
            {
                dropDown_secType.Enabled = false;
                dropDown_assetsType.Enabled = false;
            }
        }

        void ResetTypeString()
        {
            if(DataManager.IsEnable && DataManager.isWorkRange)
            {
                Worksheet sheet = GetActiveSheet();
                //当前选中的单元格
                Range range = Globals.ThisAddIn.Application.Selection;

                sheet.Cells[2, range.Column] = DataManager.GetCurrentTypeString();

                //更新UI
                UpdateDataUI();
            }
        }

        #endregion

        #region 多语言

        void LanguageInit()
        {
            LanguageManager.Init();

            comboBox_currentLanguage.Enabled = LanguageManager.IsEnable;

            button_LanguageComment.Enabled       = LanguageManager.IsEnable;
            button_deleteLanguageComment.Enabled = LanguageManager.IsEnable;
            button_openLanguageSheet.Enabled     = LanguageManager.IsEnable;
            button_changeLanguageColumn.Enabled  = LanguageManager.IsEnable;

            if (LanguageManager.IsEnable)
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

        private void comboBox_currentLanguage_TextChanged(object sender, RibbonControlEventArgs e)
        {
            LanguageManager.currentLanguage = (SystemLanguage)Enum.Parse(typeof(SystemLanguage), comboBox_currentLanguage.Text);
        }

        private void button_LanguageComment_Click(object sender, RibbonControlEventArgs e)
        {
            //只影响当前页面
            //判断当前页面是否是工作页面

            if(IsConfigWorkSheet())
            {
                Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
                //找到文件中所有多语言项
                int index = 1;
                while(!string.IsNullOrEmpty(worksheet.Cells[2,index].Value))
                {
                    string value = worksheet.Cells[2, index].Value;
                    if (value.Contains(FieldType.String + "&" + DataFieldAssetType.LocalizedLanguage))
                    {
                        //查询它们的值并写入批注
                        for (int row = 5; row <= worksheet.UsedRange.Rows.Count; row++)
                        {
                            string content = worksheet.Cells[row, index].Value;

                            if(string.IsNullOrEmpty(content))
                            {
                                continue;
                            }

                            string languageContent = LanguageManager.GetLanguageContent(LanguageManager.currentLanguage, content);

                            if (worksheet.Cells[row, index].Comment != null)
                            {
                                worksheet.Cells[row, index].Comment.Delete();
                            }

                            worksheet.Cells[row, index].AddComment(languageContent);
                        }
                    }

                    index++;
                }
            }
            else
            {
                //直接忽略
            }

        }

        private void button_deleteLanguageComment_Click(object sender, RibbonControlEventArgs e)
        {
            if (IsConfigWorkSheet())
            {
                Worksheet worksheet = Globals.ThisAddIn.Application.ActiveSheet;
                //找到文件中所有多语言项
                int index = 1;
                while (!string.IsNullOrEmpty(worksheet.Cells[2, index].Value))
                {
                    string value = worksheet.Cells[2, index].Value;
                    if (value.Contains(FieldType.String + "&" + DataFieldAssetType.LocalizedLanguage))
                    {
                        //查询它们的值并删除批注
                        for (int row = 5; row <= worksheet.UsedRange.Rows.Count; row++)
                        {
                            if (worksheet.Cells[row, index].Comment != null)
                            {
                                worksheet.Cells[row, index].Comment.Delete();
                            }
                        }
                    }

                    index++;
                }
            }
        }

        private void button_openLanguageSheet_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet configSheet = GetConfigSheet();

            //当前选中的单元格
            Range range = Globals.ThisAddIn.Application.Selection;

            string selectValue = range[1, 1].Value.ToString();

            if (string.IsNullOrEmpty(selectValue))
            {
                MessageBox.Show("没有选中有内容多语言项");
                return;
            }

            string filePath = Const.c_LanguagePrefix +"_" + LanguageManager.currentLanguage + "_"+ LanguageManager.GetFileName(selectValue);
            string sheetName = LanguageManager.GetLanguageAcronym(LanguageManager.currentLanguage) + "_" + LanguageManager.GetFileName(selectValue);

            if (string.IsNullOrEmpty(LanguageManager.GetFileName(selectValue)) && LanguageManager.CheckLanguageFileNameExist(LanguageManager.currentLanguage, sheetName))
            {
                MessageBox.Show("没有找到对应的多语言 " + sheetName);
                return;
            }

            DataConfig dataConfig;
            if (!DataConfig.IsWorkSheet(configSheet, sheetName))
            {
                //新建一条配置
                dataConfig = new DataConfig();
                dataConfig.m_txtName = filePath;
                dataConfig.m_sheetName = sheetName;

                //允许写入公式会提高效率
                dataConfig.m_coverFormula = true;

                //写入
                DataConfig.AddSheetConfig(configSheet, dataConfig);
            }
            else
            {
                dataConfig = new DataConfig(configSheet, DataConfig.GetWorkIndex(configSheet, sheetName));
            }

            if (File.Exists(dataConfig.GetTextPath()))
            {
                Worksheet wst = ExcelTool.GetSheet(Globals.ThisAddIn.Application, dataConfig.m_sheetName, true);
                DataTool.Data2Excel(dataConfig, wst);

                //激活目标页签
                wst.Activate();

                //选中目标内容
            }
        }

        private void Language_OnSheetChange(Worksheet activeSheet)
        {

        }


        void Language_OnSelectChange(Worksheet sheet, Range range)
        {
            if(LanguageManager.IsEnable)
            {
                bool isLanguage = false;
                bool isCanChangeLanguage = false;

                if(!isLanguage)
                {

                }

                button_changeLanguageColumn.Enabled = isCanChangeLanguage;
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

        Worksheet GetConfigSheet()
        {
            return ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);
        }

        Worksheet GetActiveSheet()
        {
            return Globals.ThisAddIn.Application.ActiveSheet;
        }

        bool IsConfigWorkSheet()
        {
            return DataConfig.IsWorkSheet(GetConfigSheet(), Globals.ThisAddIn.Application.ActiveSheet.Name);
        }

        void SetDropDown(RibbonDropDown dropDown,string content)
        {
            for (int i = 0; i < dropDown.Items.Count; i++)
            {
                if(dropDown.Items[i].Label == content)
                {
                    dropDown.SelectedItem = dropDown.Items[i];
                }
            }
        }

        #endregion
    }
}
