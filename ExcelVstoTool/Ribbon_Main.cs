using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelVstoTool.DialogWindow;
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

            Globals.ThisAddIn.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose; //工作簿关闭
        }


        #region 生命周期派发

        public delegate void SelectChangeHandle(Worksheet sheet, Range range);

        public static SelectChangeHandle OnSelectChange;

        private void Application_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            Data_OnWorkSheetClose();
        }

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

            //派发事件
            OnSelectChange?.Invoke(Globals.ThisAddIn.Application.ActiveSheet, Target);
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

                DataInit(config);
                UpdateDataUI();
                LanguageInit();

                SetDataUIEnabled(true);
                SetLanguageUIEnable(true);
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
                //判断文件是否存在于项目资源路径下
                if(!PathDefine.IsAssetsPath())
                {
                    MessageBox.Show("工作簿没有放在 项目 " + Const.c_DireName_Assets + " 路径下");
                    return;
                }

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
                RefreshData();
            }
            else
            {
                MessageBox.Show("导出失败");
            }
        }

        private void button_ToExcel_Click(object sender, RibbonControlEventArgs e)
        {
            //确认弹窗
            MessageBoxButtons mess = MessageBoxButtons.OKCancel;
            DialogResult dr = MessageBox.Show("确定要导入吗，此操作无法撤销", "提示", mess);
            if (dr == DialogResult.Cancel)
            {
                return;
            }

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
                        //重新生成一次枚举
                        DataManager.WriteEnumConfig(dataConfig, config);
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


        private void button_CreateNewTable_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet config = GetConfigSheet();

            //没有初始化直接返回
            if (config == null)
            {
                return;
            }

            List<string> newList = new List<string>();

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                //只有不存在的表格才会进行创建
                if (!string.IsNullOrEmpty(dataConfig.m_sheetName) 
                    && !string.IsNullOrEmpty(dataConfig.m_txtName)
                    && !dataConfig.GetFileIsExist())
                {
                    string fileName = PathDefine.GetDataPath() + @"\" + dataConfig.m_txtName + ".txt";
                    Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
                    //创建一个新表格到目标位置
                    DataTool.CreateNewData(fileName);

                    DataTool.Data2Excel(dataConfig, wst);

                    newList.Add(dataConfig.m_sheetName + "|" + dataConfig.m_txtName);
                }
            }

            if(newList.Count == 0)
            {
                MessageBox.Show("没有新表被创建");
            }
            else
            {
                string info = "创建完成";

                for (int i = 0; i < newList.Count; i++)
                {
                    info += "\n->" + newList[i];
                }

                MessageBox.Show(info);
            }
        }

        bool JudgeCanCreateTable()
        {
            Worksheet config = GetConfigSheet();

            //没有初始化直接返回
            if (config == null)
            {
                return false;
            }

            //进行转换
            for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
            {
                DataConfig dataConfig = new DataConfig(config, i);

                //只有不存在的表格才会进行创建
                if (!string.IsNullOrEmpty(dataConfig.m_sheetName)
                    && !string.IsNullOrEmpty(dataConfig.m_txtName)
                    && !dataConfig.GetFileIsExist())
                {
                    return true;
                }
            }

            return false;
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
            RefreshData();

        }

        private void button_openFile_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet configSheet = GetConfigSheet();
            string fileName = dropDown_fileList.SelectedItem.Label;
            if(!string.IsNullOrEmpty(fileName))
            {
                //判断这个页签是否存在
                Worksheet sheet = null;

                if(!GetSheetIsExist(fileName))
                {
                    sheet = GetSheet(fileName,true);

                    PerformanceSwitch(true);

                    DataConfig dataConfig = new DataConfig();
                    dataConfig.m_txtName = fileName;
                    dataConfig.m_sheetName = fileName;

                    //写入
                    DataConfig.AddSheetConfig(configSheet, dataConfig);

                    if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
                    {
                        if (File.Exists(dataConfig.GetTextPath()))
                        {
                            Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
                            DataTool.Data2Excel(dataConfig, wst);
                            //重新生成一次枚举
                            DataManager.WriteEnumConfig(dataConfig, configSheet);
                        }
                        else
                        {
                            //TODO 打开失败反馈
                            MessageBox.Show("打开失败 文件不存在 " + dataConfig.m_txtName);
                        }
                    }

                    PerformanceSwitch(false);
                }
                else
                {
                    sheet = GetSheet(fileName);
                }

                sheet.Activate();
            }
        }

        private void button_closeFile_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet config = GetConfigSheet();
            DataConfig dataConfig = GetActiveDataConfig();
            Worksheet worksheet = GetActiveSheet();

            //MessageBoxButtons mess = MessageBoxButtons.OKCancel;
            //DialogResult dr = MessageBox.Show("确定要关闭 " + dataConfig.m_sheetName + "|" + dataConfig.m_txtName, "提示", mess);
            //if (dr == DialogResult.Cancel)
            //{
            //    return;
            //}

            dataConfig.Delete(config);
            worksheet.Delete();
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

        private void button_deleteTable_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet config = GetConfigSheet();
            DataConfig dataConfig = GetActiveDataConfig();

            //确认弹窗
            MessageBoxButtons mess = MessageBoxButtons.OKCancel;
            DialogResult dr = MessageBox.Show("确定要删除 " + dataConfig.m_sheetName + "|" + dataConfig.m_txtName, "提示",mess);
            if(dr == DialogResult.Cancel)
            {
                return;
            }

            //删除页面
            Worksheet wsheet = GetActiveSheet();
            wsheet.Delete();

            //删除文件
            if(File.Exists(dataConfig.GetTextPath()))
            {
                File.Delete(dataConfig.GetTextPath());
            }
            else
            {
                MessageBox.Show("没有找到 " + dataConfig.GetTextPath() +"");
            }

            //删除配置
            dataConfig.Delete(config);
        }

        void GenerateDropDownList(int totalRow, int col,FieldType fieldType, DataFieldAssetType assetType,string secType)
        {
            //构造下拉列表
            List<String> list = new List<string>();
            if (fieldType == FieldType.Enum)
            {
                list = DataManager.GetEnumList(secType);
            }
            else if (fieldType == FieldType.Bool)
            {
                list.Add("TRUE");
                list.Add("FALSE");
            }
            else if (fieldType == FieldType.String)
            {
                if (assetType == DataFieldAssetType.Texture)
                {
                    list = DataManager.GetTextureList();
                }
                else if (assetType == DataFieldAssetType.Prefab)
                {
                    list = DataManager.GetPrefabList();
                }
                else if (assetType == DataFieldAssetType.TableName)
                {
                    list = DataManager.TableName;
                }
                else if (assetType == DataFieldAssetType.TableKey)
                {
                    list = DataManager.GetTableKeyList(DataManager.CurrentSecType);
                }
                else if (assetType == DataFieldAssetType.LocalizedLanguage)
                {
                    if (secType != "")
                    {
                        list = LanguageManager.GetLanguageKeyList(LanguageManager.currentLanguage, secType);
                    }
                }
            }

            //只保留前500的数据，避免报错
            if(list.Count > 200)
            {
                list.RemoveRange(200, list.Count - 200);
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

        private void button_importSingleTable_Click(object sender, RibbonControlEventArgs e)
        {
            //确认弹窗
            MessageBoxButtons mess = MessageBoxButtons.OKCancel;
            DialogResult dr = MessageBox.Show("确定要导入吗，此操作无法撤销", "提示", mess);
            if (dr == DialogResult.Cancel)
            {
                return;
            }

            DataConfig dataConfig = GetActiveDataConfig();
            Worksheet wst = GetSheet(dataConfig.m_sheetName);

            DateTime now = System.DateTime.Now;

            PerformanceSwitch(true);

            //进行转换
            if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
            {
                if (File.Exists(dataConfig.GetTextPath()))
                {
                    DataTool.Data2Excel(dataConfig, wst);
                    //重新生成一次枚举
                    DataManager.WriteEnumConfig(dataConfig, GetConfigSheet());
                }
                else
                {
                    MessageBox.Show("没有找到文件 " + dataConfig.GetTextPath());
                }
            }

            PerformanceSwitch(false);

            MessageBox.Show("导入完成\n用时：" + (DateTime.Now - now).TotalSeconds + "s");
        }

        private void button_exportSingleTable_Click(object sender, RibbonControlEventArgs e)
        {
            DataConfig dataConfig = GetActiveDataConfig();
            Worksheet wst = GetSheet(dataConfig.m_sheetName);

            if (!string.IsNullOrEmpty(dataConfig.m_sheetName) 
                && !string.IsNullOrEmpty(dataConfig.GetTextPath())
                && wst != null)
            {
                DataTable dataTable = null;

                //为了提高导出效率，所以尽量减少excel的读取次数
                //将检验过后的DataTable直接进行序列化
                if (checkBox_exportCheck.Checked)
                {
                    dataTable = CheckTool.CheckSheet(wst, dataConfig);
                    if (dataTable != null)
                    {
                        FileTool.WriteStringByFile(dataConfig.GetTextPath(), DataTable.Serialize(dataTable));
                    }
                    else
                    {
                        MessageBox.Show("导出失败：校验未能通过: " + dataConfig.m_sheetName + "|" + dataConfig.m_txtName);
                        return;
                    }
                }
                else
                {
                    dataTable = DataTool.Excel2Table(wst, dataConfig);
                    FileTool.WriteStringByFile(dataConfig.GetTextPath(), DataTable.Serialize(dataTable));

                    RefreshData();
                }

                MessageBox.Show("导出完毕");
            }
            else
            {
                MessageBox.Show("导出失败：配置或者文件不正确 " + dataConfig.m_sheetName);
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
            UpdateLanguageUI();
        }

        private void dropDown_assetsType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            DataManager.CurrentAssetType = (DataFieldAssetType)Enum.Parse(typeof(DataFieldAssetType), dropDown_assetsType.SelectedItem.Label);

            ResetTypeString();
            UpdateLanguageUI();
        }

        private void dropDown_secType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            DataManager.CurrentSecType = dropDown_secType.SelectedItem.Label;
            ResetTypeString();
            UpdateLanguageUI();
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

        private void Data_OnWorkSheetClose()
        {
            if(DataManager.IsEnable)
            {
                //退出时清除所有的数据校验，以免出错
                Worksheet config = GetConfigSheet();

                //没有初始化直接返回
                if (config == null)
                {
                    return;
                }

                for (int i = 2; i < config.UsedRange.Rows.Count + 1; i++)
                {
                    DataConfig dataConfig = new DataConfig(config, i);

                    if (!string.IsNullOrEmpty(dataConfig.m_sheetName))
                    {
                        Worksheet wst = GetSheet(dataConfig.m_sheetName, false);
                        
                        if(wst != null)
                        {
                            wst.UsedRange.Validation.Delete();
                        }
                    }
                }

                //保存
                Globals.ThisAddIn.Application.ActiveWorkbook.Save();
            }
        }

        #endregion

        #endregion

        #region UI更新逻辑

        void UpdateFileList()
        {
            dropDown_fileList.Items.Clear();

            List<string> list = DataManager.TableName;
            for (int i = 0; i < list.Count; i++)
            {
                RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                tmp.Label = list[i];
                dropDown_fileList.Items.Add(tmp);
            }
        }

        void UpdateDataUI()
        {
            //文件操作
            UpdateFileList();
            button_closeFile.Enabled = IsConfigWorkSheet();

            button_generateDataClass.Enabled = IsConfigWorkSheet();
            button_CreateDataDropDownList.Enabled = IsConfigWorkSheet();
            button_ClearDropDownList.Enabled = IsConfigWorkSheet();
            button_deleteTable.Enabled = IsConfigWorkSheet();
            button_importSingleTable.Enabled = IsConfigWorkSheet();
            button_exportSingleTable.Enabled = IsConfigWorkSheet();

            button_createNewTable.Enabled = JudgeCanCreateTable();

            ResetSecTypeDropDownItem();

            SetDropDown(dropDown_dataType, DataManager.CurrentFieldType.ToString());
            SetDropDown(dropDown_assetsType, DataManager.CurrentAssetType.ToString());
            SetDropDown(dropDown_secType, DataManager.CurrentSecType.ToString());

            if (!DataManager.isWorkRange || !IsConfigWorkSheet())
            {
                dropDown_dataType.Enabled = false;
                dropDown_secType.Enabled = false;
                dropDown_assetsType.Enabled = false;

                return;
            }

            dropDown_dataType.Enabled = true;
            if (DataManager.CurrentFieldType == FieldType.String 
                || DataManager.CurrentFieldType == FieldType.StringArray)
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
            else if (DataManager.CurrentFieldType == FieldType.Enum 
                || DataManager.CurrentFieldType == FieldType.EnumArray)
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

            //文件操作
            dropDown_fileList.Enabled = isEnable;
            button_openFile.Enabled = isEnable;
            button_closeFile.Enabled = isEnable;

            button_createNewTable.Enabled = isEnable;
            button_ToExcel.Enabled = isEnable;
            button_toTxt.Enabled = isEnable;
            button_importSingleTable.Enabled = isEnable;
            button_exportSingleTable.Enabled = isEnable;

            button_check.Enabled = isEnable;
            button_CreateDataDropDownList.Enabled = isEnable;
            button_ClearDropDownList.Enabled = isEnable;
            button_generateDataClass.Enabled = isEnable;
            button_deleteTable.Enabled = isEnable;
            
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
            else if (DataManager.CurrentFieldType == FieldType.Enum 
                || DataManager.CurrentFieldType == FieldType.EnumArray)
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

            UpdateLanguageUI();
        }

        #region 事件监听

        #region UI交互事件

        private void comboBox_currentLanguage_TextChanged(object sender, RibbonControlEventArgs e)
        {
            LanguageManager.currentLanguage = (SystemLanguage)Enum.Parse(typeof(SystemLanguage), comboBox_currentLanguage.Text);
        }

        private void button_openLanguageFile_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet configSheet = GetConfigSheet();
            string fileName = dropDown_languageFileList.SelectedItem.Label;

            string filePath = Const.c_LanguagePrefix + "_" + LanguageManager.currentLanguage + "_" + fileName;
            string sheetName = CutSheetName( LanguageManager.GetLanguageAcronym(LanguageManager.currentLanguage) + "_" + fileName);

            if (!string.IsNullOrEmpty(fileName))
            {
                //判断这个页签是否存在
                Worksheet sheet = null;

                if (!GetSheetIsExist(sheetName))
                {
                    sheet = GetSheet(sheetName, true);

                    PerformanceSwitch(true);

                    DataConfig dataConfig = new DataConfig();
                    dataConfig.m_txtName = filePath;
                    dataConfig.m_sheetName = sheetName;

                    //写入
                    DataConfig.AddSheetConfig(configSheet, dataConfig);

                    if (!string.IsNullOrEmpty(dataConfig.m_sheetName) && !string.IsNullOrEmpty(dataConfig.GetTextPath()))
                    {
                        if (File.Exists(dataConfig.GetTextPath()))
                        {
                            Worksheet wst = GetSheet(dataConfig.m_sheetName, true);
                            DataTool.Data2Excel(dataConfig, wst);
                        }
                        else
                        {
                            MessageBox.Show("打开失败 文件不存在 " + dataConfig.m_txtName);
                        }
                    }

                    PerformanceSwitch(false);
                }
                else
                {
                    sheet = GetSheet(sheetName);
                }

                sheet.Activate();
            }
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
                    FieldTypeStruct typeStruct = DataManager.PaseToFieldStructType(value);
                    if (typeStruct.fieldType == FieldType.String 
                        && typeStruct.assetType == DataFieldAssetType.LocalizedLanguage)
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
                    FieldTypeStruct typeStruct = DataManager.PaseToFieldStructType(value);
                    if (typeStruct.fieldType == FieldType.String
                        && typeStruct.assetType == DataFieldAssetType.LocalizedLanguage)
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
            string sheetName = CutSheetName(LanguageManager.GetLanguageAcronym(LanguageManager.currentLanguage) + "_" + LanguageManager.GetFileName(selectValue));
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
            
            Worksheet sheet = GetActiveSheet();
            DataConfig config = GetActiveDataConfig();

            string fileName = config.m_txtName + "_" + GetCurrentKeyName();
            int col = GetCurrentSelectRange().Column;
            int row = 5;

            //进行重名判断

            if(LanguageManager.CheckLanguageFileNameExist(LanguageManager.currentLanguage, fileName))
            {
                MessageBox.Show(fileName + " 文件名已存在");
                return;
            }

            //进行弹窗确认
            MessageBoxButtons mess = MessageBoxButtons.OKCancel;
            DialogResult dr = MessageBox.Show("确定要以 " + fileName + " 创建多语言文件吗？","提示", mess);

            if(dr != DialogResult.OK)
            {
                return;
            }

            //构造语言数据
            Dictionary<string, string> languageData = new Dictionary<string, string>();
            while (!string.IsNullOrEmpty( sheet.Cells[row, 1].Text))
            {
                string key = sheet.Cells[row, 1].Text;
                string value = sheet.Cells[row, col].Text;

                languageData.Add(key, value);

                //修改现有的表格
                sheet.Cells[row, col].Value = (fileName + "_" + key).Replace("_", "/");

                row++;
            }

            //构造新的多语言文件
            LanguageManager.CreateLanguageFile(LanguageManager.currentLanguage, fileName, languageData);

            //修改表头
            DataManager.CurrentAssetType = DataFieldAssetType.LocalizedLanguage;
            DataManager.CurrentSecType = fileName;

            ResetTypeString();


            MessageBox.Show("转换完毕");
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

        void UpdateLanguageFileList()
        {
            dropDown_languageFileList.Items.Clear();

            List<string> list = LanguageManager.GetLanguageFileNameList(LanguageManager.currentLanguage);
            for (int i = 0; i < list.Count; i++)
            {
                RibbonDropDownItem tmp = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();

                tmp.Label = list[i];
                dropDown_languageFileList.Items.Add(tmp);
            }
        }

        void UpdateLanguageUI()
        {
            //依赖DataManager的解析
            if (DataManager.IsEnable && LanguageManager.IsEnable)
            {
                UpdateLanguageFileList();

                //comboBox_currentLanguage.Enabled = IsConfigWorkSheet();
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
            dropDown_languageFileList.Enabled = LanguageManager.IsEnable;
            button_openLanguageFile.Enabled = LanguageManager.IsEnable;

            button_LanguageComment.Enabled = LanguageManager.IsEnable;
            button_deleteLanguageComment.Enabled = LanguageManager.IsEnable;

            button_openLanguageSheet.Enabled = LanguageManager.IsEnable;
            button_changeLanguageColumn.Enabled = LanguageManager.IsEnable;

            comboBox_currentLanguage.Enabled = LanguageManager.IsEnable;
        }

        #endregion

        #endregion

        #region 数据工具

        private void button_ArraryToolWindow_Click(object sender, RibbonControlEventArgs e)
        {
            ArrayToolWindow atw = new ArrayToolWindow();
            atw.Show();
            //将当前选中区域传入
            atw.OnSelectChange(GetActiveSheet(), GetCurrentSelectRange());
        }

        private void button_CompressToolWindow_Click(object sender, RibbonControlEventArgs e)
        {
            CompressToolWindow ctw = new CompressToolWindow();
            ctw.Show();
            //将当前选中区域传入
            ctw.OnSelectChange(GetActiveSheet(), GetCurrentSelectRange());
        }

        #endregion

        #region 工具方法

        void RefreshData()
        {
            DataInit(GetConfigSheet());
            LanguageInit();
        }

        //性能开关
        void PerformanceSwitch(bool enable)
        {
            if (checkBox_CloseView.Checked)
            {
                Globals.ThisAddIn.Application.Visible = !enable;
                Globals.ThisAddIn.Application.ScreenUpdating = !enable;
                Globals.ThisAddIn.Application.EnableEvents = !enable;
            }
        }

        public static Worksheet GetSheet(string shetName, bool isCreate = false)
        {
            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, Const.c_SheetName_Config))
            {
                return null;
            }

            return ExcelTool.GetSheet(Globals.ThisAddIn.Application, shetName, isCreate);
        }

        public static bool GetSheetIsExist(string sheetName)
        {
            return ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, sheetName);
        }


        Worksheet GetConfigSheet()
        {
            if (!ExcelTool.ExistSheetName(Globals.ThisAddIn.Application, Const.c_SheetName_Config))
            {
                return null;
            }

            return ExcelTool.GetSheet(Globals.ThisAddIn.Application, Const.c_SheetName_Config);
        }


        public static Worksheet GetActiveSheet()
        {
            return Globals.ThisAddIn.Application.ActiveSheet;
        }

        public static Range GetCurrentSelectRange()
        {
            //当前选中的单元格
            return  Globals.ThisAddIn.Application.Selection;
        }

        string GetCurrentKeyName()
        {
            Range range = GetCurrentSelectRange();

            return GetActiveSheet().Cells[1, range.Column].Text;
        }

        /// <summary>
        /// 判断当前页面是否是被config管理的页面
        /// </summary>
        /// <returns></returns>
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

        string CutSheetName(string sheetName)
        {
            if(sheetName.Length < 30)
            {
                return sheetName;
            }
            else
            {
                return sheetName.Substring(0, 30);
            }
        }

        public static Range GetRangeByRangeString(string rangeString)
        {
            string SheetName = rangeString.Split('!')[0];
            string range = rangeString.Split('!')[1];

            Worksheet worksheet = GetSheet(SheetName, false);

            return worksheet.Range[range];
        }

        #endregion


    }
}
