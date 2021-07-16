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

        #region 生命周期派发

        private void Application_WorkbookActivate(Workbook Wb)
        {
            ConfigLogic();
        }

        private void Application_SheetActivate(object Sh)
        {
            ConfigLogic();
            Data_OnSheetChange(Globals.ThisAddIn.Application.ActiveSheet);
            Language_OnSheetChange(Globals.ThisAddIn.Application.ActiveSheet);
        }

        private void Application_SheetSelectionChange(object Sh, Range Target)
        {
            ConfigLogic();
            Data_OnSelectChange(Globals.ThisAddIn.Application.ActiveSheet, Target);
            Language_OnSelectChange(Globals.ThisAddIn.Application.ActiveSheet, Target);
        }

        #endregion

        #region 设置

        void ConfigLogic()
        {
            //如果读取到设置页签则自动进行初始化
            Worksheet config = GetConfigSheet();
            if (config != null)
            {
                //隐藏初始化按钮
                button_initConfig.Visible = false;

                SetDataUIEnabled(true);
                SetLanguageUIEnable(true);

                if (!DataManager.IsEnable)
                {
                    DataInit(config);
                    UpdateDataUI();
                }

                if (!LanguageManager.IsEnable)
                {
                    LanguageInit();
                }
            }
            else
            {
                button_initConfig.Visible = true;
                DataManager.IsEnable = false;
                LanguageManager.IsEnable = false;

                SetDataUIEnabled(false);
                SetLanguageUIEnable(false);
            }
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

        #endregion

        #region 导入导出

        private void button_toTxt_Click(object sender, RibbonControlEventArgs e)
        {
            //先进行一次保存
            Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            Worksheet config = GetConfigSheet();

            //刷新一次数据
            DataManager.Init(GetConfigSheet());
            LanguageManager.Init();

            //没有初始化直接返回
            if (config == null)
            {
                return;
            }

            DateTime now = System.DateTime.Now;

            bool allSuccess = true;
            Dictionary<string, DataTable> result = new Dictionary<string, DataTable>();

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
                {
                    Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
                    DataTable dataTable = null;

                    //为了提高导出效率，所以尽量减少excel的读取次数
                    //将检验过后的DataTable直接进行序列化
                    if (checkBox_exportCheck.Checked)
                    {
                        dataTable = CheckTool.CheckSheet(wst, dataConfig);
                        if (dataTable == null)
                        {
                            allSuccess = false;
                            break;
                        }
                    }
                    else
                    {
                        dataTable = DataTool.Excel2Table(wst, dataConfig);
                    }

                    result.Add(dataConfig.GetTextPath(), dataTable);
                }
            }

            if (allSuccess)
            {
                foreach (var item in result)
                {
                    FileTool.WriteStringByFile(item.Key, DataTable.Serialize(item.Value));
                }
                MessageBox.Show("导出完毕\n用时：" + (DateTime.Now - now).TotalSeconds + "s");
            }
            else
            {
                MessageBox.Show("导出失败");
            }
        }

        private void button_ToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            string info = "导入完毕";
            List<string> nofindPath = new List<string>();

            Worksheet config = GetConfigSheet();

            //没有初始化直接返回
            if (config == null)
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
                    if (File.Exists(dataConfig.GetTextPath()))
                    {
                        Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
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

        void DataInit(Worksheet config)
        {
            DataManager.Init(config);

            //构造类型的下拉列表
            dropDown_dataType.Items.Clear();
            foreach (FieldType fieldType in Enum.GetValues(typeof(FieldType)))
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
        }

        #region 事件监听

        #region UI交互事件

        private void button_dataInit_Click(object sender, RibbonControlEventArgs e)
        {
            DataInit(GetConfigSheet());
            LanguageInit();
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
                    Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
                    if (CheckTool.CheckSheet(wst, dataConfig) == null)
                    {
                        info += "\n->" + dataConfig.m_sheetName + " 校验未能通过";
                    }
                }
            }

            //构造输出信息
            info += "\n用时：" + (DateTime.Now - now).TotalSeconds + "s";

            MessageBox.Show(info);
        }

        private void button_generateDataClass_Click(object sender, RibbonControlEventArgs e)
        {
            DataConfig dataConfig = GetActiveDataConfig();
            DataTable table = null;

            DateTime now = System.DateTime.Now;

            if (checkBox_exportCheck.Checked)
            {
                table = CheckTool.CheckSheet(GetActiveSheet(), dataConfig);
            }
            else
            {
                table = DataTool.Excel2Table(GetActiveSheet(), dataConfig);
            }

            if (table != null)
            {
                string dataName = dataConfig.m_txtName;
                string csPath = PathDefine.GetDataGeneratePath() + @"\" + dataName + "Generate.cs";
                string content = DataTool.CreateDataCSharpFile(dataName, table);

                FileTool.WriteStringByFile(csPath, content);

                MessageBox.Show("生成完毕\n用时：" + (DateTime.Now - now).TotalSeconds + "s");
            }
            else
            {
                MessageBox.Show("生成失败");
            }
        }

        private void button_CreateDataDropDownList_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet sheet = GetActiveSheet();
            int totalRow = sheet.UsedRange.Rows.Count;
            int col = 2;
            int row = 2;

            while (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
            {
                string typeString = sheet.Cells[row, col].Text;
                FieldTypeStruct typeStruct = DataManager.PaseToFieldStructType(typeString);
                GenerateDropDownList(totalRow, col, typeStruct.fieldType, typeStruct.assetType, typeStruct.secType);

                col++;
            }
        }

        void GenerateDropDownList(int totalRow, int col,FieldType fieldType, DataFieldAssetType assetType,string secType)
        {
            //构造下拉列表
            List<String> list = new List<string>();
            if (DataManager.CurrentFieldType == FieldType.Enum)
            {
                list = DataManager.GetEnumList(DataManager.CurrentSecType);
            }
            else if (DataManager.CurrentFieldType == FieldType.String)
            {
                if (DataManager.CurrentAssetType == DataFieldAssetType.Texture)
                {
                    //数据量过大，暂不支持生成下拉列表
                    //list = DataManager.GetTextureList();
                }
                else if (DataManager.CurrentAssetType == DataFieldAssetType.Prefab)
                {
                    //数据量过大，暂不支持生成下拉列表
                    //list = DataManager.GetPrefabList();
                }
                else if (DataManager.CurrentAssetType == DataFieldAssetType.TableName)
                {
                    list = DataManager.TableName;
                }
                else if (DataManager.CurrentAssetType == DataFieldAssetType.TableKey)
                {
                    list = DataManager.GetTableKeyList(DataManager.CurrentSecType);
                }
                else if (DataManager.CurrentAssetType == DataFieldAssetType.LocalizedLanguage)
                {
                    if (DataManager.CurrentSecType != "")
                    {
                        list = LanguageManager.GetLanguageKeyList(LanguageManager.currentLanguage, DataManager.CurrentSecType);
                    }
                }
            }

            //确定下拉范围
            string colName = ExcelTool.Int2ColumnName(col);
            Range aimRange = GetActiveSheet().Range[colName + "4:" + colName + totalRow];

            if (list.Count > 0)
            {
                string values = string.Join(",", list);
                aimRange.Validation.Delete();
                aimRange.Validation.Add(
                    Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                    Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop,
                    Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween,
                    values,
                    Type.Missing);
            }
            else
            {
                aimRange.Validation.Delete();
            }
        }

        private void button_ClearDropDownList_Click(object sender, RibbonControlEventArgs e)
        {
            GetActiveSheet().UsedRange.Validation.Delete();
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

        private void dropDown_secType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            DataManager.CurrentSecType = dropDown_secType.SelectedItem.Label;
            ResetTypeString();
        }

        #endregion

        #region 生命周期回调

        private void Data_OnSheetChange(Worksheet activeSheet)
        {
            if (DataManager.IsEnable)
            {
                //更新UI
                UpdateDataUI();
            }
        }

        private void Data_OnSelectChange(Worksheet sheet, Range target)
        {
            if (DataManager.IsEnable)
            {
                if (IsConfigWorkSheet())
                {
                    //第一列不处理
                    if (target.Column == 1)
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

        #endregion

        #endregion

        #region UI更新逻辑

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
                else if (DataManager.CurrentAssetType == DataFieldAssetType.LocalizedLanguage)
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

        private void SetDataUIEnabled(bool isEnable)
        {
            button_refreshData.Enabled = isEnable;

            button_ToExcel.Enabled = isEnable;
            button_toTxt.Enabled = isEnable;

            button_check.Enabled = isEnable;
            button_CreateDataDropDownList.Enabled = isEnable;
            button_generateDataClass.Enabled = isEnable;

            dropDown_assetsType.Enabled = isEnable;
            dropDown_dataType.Enabled = isEnable;
            dropDown_secType.Enabled = isEnable;

            UpdateDataUI();
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
            if (DataManager.CurrentAssetType == DataFieldAssetType.LocalizedLanguage)
            {
                List<string> list = LanguageManager.GetLanguageFileNameList(LanguageManager.currentLanguage);

                for (int i = 0; i < list.Count; i++)
                {
                    RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                    tmp.Label = list[i];
                    dropDown_secType.Items.Add(tmp);
                }
            }
            else if (DataManager.CurrentFieldType == FieldType.Enum)
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

        void ResetTypeString()
        {
            if (DataManager.IsEnable && DataManager.isWorkRange)
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

        #endregion

        #region 多语言

        void LanguageInit()
        {
            LanguageManager.Init();

            comboBox_currentLanguage.Enabled = LanguageManager.IsEnable;

            button_LanguageComment.Enabled = LanguageManager.IsEnable;
            button_deleteLanguageComment.Enabled = LanguageManager.IsEnable;
            button_openLanguageSheet.Enabled = LanguageManager.IsEnable;
            button_changeLanguageColumn.Enabled = LanguageManager.IsEnable;

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

        #region 事件监听

        #region UI交互事件

        private void comboBox_currentLanguage_TextChanged(object sender, RibbonControlEventArgs e)
        {
            LanguageManager.currentLanguage = (SystemLanguage)Enum.Parse(typeof(SystemLanguage), comboBox_currentLanguage.Text);
        }

        private void button_LanguageComment_Click(object sender, RibbonControlEventArgs e)
        {
            //只影响当前页面
            //判断当前页面是否是工作页面

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
                        //查询它们的值并写入批注
                        for (int row = 5; row <= worksheet.UsedRange.Rows.Count; row++)
                        {
                            string content = worksheet.Cells[row, index].Value;

                            if (string.IsNullOrEmpty(content))
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

            string filePath = Const.c_LanguagePrefix + "_" + LanguageManager.currentLanguage + "_" + LanguageManager.GetFileName(selectValue);
            string sheetName = LanguageManager.GetLanguageAcronym(LanguageManager.currentLanguage) + "_" + LanguageManager.GetFileName(selectValue);
            string key = LanguageManager.GetLanguageKey(selectValue);

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
                Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
                DataTool.Data2Excel(dataConfig, wst);

                //激活目标页签
                wst.Activate();

                int row = 1;
                //选中目标内容
                while (!string.IsNullOrEmpty(wst.Cells[row, 1].Text))
                {
                    string value = wst.Cells[row, 1].Text;
                    if (value  == key)
                    {
                        wst.Cells[row, 2].Select();
                        break;
                    }

                    row++;
                }
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

        #region 生命周期

        private void Language_OnSheetChange(Worksheet activeSheet)
        {
            UpdateLanguageUI();
        }

        void Language_OnSelectChange(Worksheet sheet, Range range)
        {
            UpdateLanguageUI();
        }


        #endregion

        #endregion

        #region UI更新逻辑

        void UpdateLanguageUI()
        {
            //依赖DataManager的解析
            if (DataManager.IsEnable && LanguageManager.IsEnable)
            {
                button_LanguageComment.Enabled = IsConfigWorkSheet();
                button_deleteLanguageComment.Enabled = IsConfigWorkSheet();

                bool isCanChangeLanguage = false;
                bool isCanOpenLanguage = false;

                if (DataManager.CurrentFieldType == FieldType.String
                    && IsConfigWorkSheet()
                    && DataManager.isWorkRange)
                {
                    if (DataManager.CurrentAssetType == DataFieldAssetType.LocalizedLanguage)
                    {
                        isCanOpenLanguage = true;
                    }
                    else
                    {
                        isCanChangeLanguage = true;
                    }
                }

                button_openLanguageSheet.Enabled = isCanOpenLanguage;
                button_changeLanguageColumn.Enabled = isCanChangeLanguage;
            }
        }


        private void SetLanguageUIEnable(bool isEnable)
        {
            button_LanguageComment.Enabled = isEnable;
            button_deleteLanguageComment.Enabled = isEnable;

            button_openLanguageSheet.Enabled = isEnable;
            button_changeLanguageColumn.Enabled = isEnable;

            comboBox_currentLanguage.Enabled = isEnable;
        }

        #endregion

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

        Worksheet GetSheet(string shetName, bool isCreate = false)
        {
            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, Const.c_SheetName_Config))
            {
                return null;
            }

            return ExcelTool.GetSheet(Globals.ThisAddIn.Application, shetName, isCreate);
        }


        Worksheet GetConfigSheet()
        {
            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, Const.c_SheetName_Config))
            {
                return null;
            }

            return ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);
        }


        Worksheet GetActiveSheet()
        {
            return Globals.ThisAddIn.Application.ActiveSheet;
        }

        Range GetCurrentSelectRange()
        {
            //当前选中的单元格
            return  Globals.ThisAddIn.Application.Selection;
        }

        bool IsConfigWorkSheet()
        {
            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, Const.c_SheetName_Config))
            {
                return false;
            }

            return DataConfig.IsWorkSheet(GetConfigSheet(), Globals.ThisAddIn.Application.ActiveSheet.Name);
        }

        DataConfig GetActiveDataConfig()
        {
            Worksheet config = GetConfigSheet();
            return new DataConfig(config, DataConfig.GetWorkIndex(config, Globals.ThisAddIn.Application.ActiveSheet.Name));
        }

        void SetDropDown(RibbonDropDown dropDown, string content)
        {
            for (int i = 0; i < dropDown.Items.Count; i++)
            {
                if (dropDown.Items[i].Label == content)
                {
                    dropDown.SelectedItem = dropDown.Items[i];
                }
            }
        }

        #endregion


    }
}
