namespace ExcelVstoTool
{
    partial class Ribbon_Main : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon_Main()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab_main = this.Factory.CreateRibbonTab();
            this.group_main = this.Factory.CreateRibbonGroup();
            this.checkBox_CloseView = this.Factory.CreateRibbonCheckBox();
            this.checkBox_exportCheck = this.Factory.CreateRibbonCheckBox();
            this.group_InOut = this.Factory.CreateRibbonGroup();
            this.group_Data = this.Factory.CreateRibbonGroup();
            this.group_language = this.Factory.CreateRibbonGroup();
            this.comboBox_currentLanguage = this.Factory.CreateRibbonComboBox();
            this.button_initData = this.Factory.CreateRibbonButton();
            this.button_ToExcel = this.Factory.CreateRibbonButton();
            this.button_toTxt = this.Factory.CreateRibbonButton();
            this.button_dataInit = this.Factory.CreateRibbonButton();
            this.button_check = this.Factory.CreateRibbonButton();
            this.button_changeLanguageColumn = this.Factory.CreateRibbonButton();
            this.button_LanguageInfo = this.Factory.CreateRibbonButton();
            this.tab_main.SuspendLayout();
            this.group_main.SuspendLayout();
            this.group_InOut.SuspendLayout();
            this.group_Data.SuspendLayout();
            this.group_language.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_main
            // 
            this.tab_main.Groups.Add(this.group_main);
            this.tab_main.Groups.Add(this.group_InOut);
            this.tab_main.Groups.Add(this.group_Data);
            this.tab_main.Groups.Add(this.group_language);
            this.tab_main.Label = "拓展工具";
            this.tab_main.Name = "tab_main";
            this.tab_main.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            // 
            // group_main
            // 
            this.group_main.Items.Add(this.button_initData);
            this.group_main.Items.Add(this.checkBox_CloseView);
            this.group_main.Items.Add(this.checkBox_exportCheck);
            this.group_main.Label = "设置";
            this.group_main.Name = "group_main";
            // 
            // checkBox_CloseView
            // 
            this.checkBox_CloseView.Checked = true;
            this.checkBox_CloseView.Label = "关闭窗口以提高执行效率";
            this.checkBox_CloseView.Name = "checkBox_CloseView";
            // 
            // checkBox_exportCheck
            // 
            this.checkBox_exportCheck.Checked = true;
            this.checkBox_exportCheck.Label = "导出时检验数据有效性";
            this.checkBox_exportCheck.Name = "checkBox_exportCheck";
            // 
            // group_InOut
            // 
            this.group_InOut.Items.Add(this.button_ToExcel);
            this.group_InOut.Items.Add(this.button_toTxt);
            this.group_InOut.Label = "导入导出";
            this.group_InOut.Name = "group_InOut";
            // 
            // group_Data
            // 
            this.group_Data.Items.Add(this.button_dataInit);
            this.group_Data.Items.Add(this.button_check);
            this.group_Data.Label = "数据";
            this.group_Data.Name = "group_Data";
            // 
            // group_language
            // 
            this.group_language.Items.Add(this.comboBox_currentLanguage);
            this.group_language.Items.Add(this.button_changeLanguageColumn);
            this.group_language.Items.Add(this.button_LanguageInfo);
            this.group_language.Label = "多语言";
            this.group_language.Name = "group_language";
            // 
            // comboBox_currentLanguage
            // 
            this.comboBox_currentLanguage.Enabled = false;
            this.comboBox_currentLanguage.Label = "当前语言";
            this.comboBox_currentLanguage.Name = "comboBox_currentLanguage";
            this.comboBox_currentLanguage.ShowItemImage = false;
            this.comboBox_currentLanguage.Text = null;
            // 
            // button_initData
            // 
            this.button_initData.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_initData.Label = "设置初始化";
            this.button_initData.Name = "button_initData";
            this.button_initData.OfficeImageId = "AdpDiagramNewTable";
            this.button_initData.ShowImage = true;
            this.button_initData.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_initData_Click);
            // 
            // button_ToExcel
            // 
            this.button_ToExcel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_ToExcel.Label = "从文本导入";
            this.button_ToExcel.Name = "button_ToExcel";
            this.button_ToExcel.OfficeImageId = "ImportTextFile";
            this.button_ToExcel.ShowImage = true;
            this.button_ToExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_ToExcel_Click);
            // 
            // button_toTxt
            // 
            this.button_toTxt.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_toTxt.Label = "导出到文本";
            this.button_toTxt.Name = "button_toTxt";
            this.button_toTxt.OfficeImageId = "ExportTextFile";
            this.button_toTxt.ShowImage = true;
            this.button_toTxt.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_toTxt_Click);
            // 
            // button_dataInit
            // 
            this.button_dataInit.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_dataInit.Label = "外部数据初始化";
            this.button_dataInit.Name = "button_dataInit";
            this.button_dataInit.OfficeImageId = "FileCompactAndRepairDatabase";
            this.button_dataInit.ShowImage = true;
            this.button_dataInit.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_dataInit_Click);
            // 
            // button_check
            // 
            this.button_check.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_check.Label = "数据校验";
            this.button_check.Name = "button_check";
            this.button_check.OfficeImageId = "AcceptInvitation";
            this.button_check.ShowImage = true;
            this.button_check.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_check_Click);
            // 
            // button_changeLanguageColumn
            // 
            this.button_changeLanguageColumn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button_changeLanguageColumn.Enabled = false;
            this.button_changeLanguageColumn.Label = "转化为多语言列";
            this.button_changeLanguageColumn.Name = "button_changeLanguageColumn";
            this.button_changeLanguageColumn.OfficeImageId = "SetLanguage";
            this.button_changeLanguageColumn.ShowImage = true;
            this.button_changeLanguageColumn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_changeLanguageColumn_Click);
            // 
            // button_LanguageInfo
            // 
            this.button_LanguageInfo.Label = "帮助";
            this.button_LanguageInfo.Name = "button_LanguageInfo";
            this.button_LanguageInfo.OfficeImageId = "Help";
            this.button_LanguageInfo.ShowImage = true;
            this.button_LanguageInfo.ShowLabel = false;
            this.button_LanguageInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button_LanguageInfo_Click);
            // 
            // Ribbon_Main
            // 
            this.Name = "Ribbon_Main";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab_main);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Main_Load);
            this.tab_main.ResumeLayout(false);
            this.tab_main.PerformLayout();
            this.group_main.ResumeLayout(false);
            this.group_main.PerformLayout();
            this.group_InOut.ResumeLayout(false);
            this.group_InOut.PerformLayout();
            this.group_Data.ResumeLayout(false);
            this.group_Data.PerformLayout();
            this.group_language.ResumeLayout(false);
            this.group_language.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_main;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_main;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_InOut;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_toTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_ToExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_initData;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_CloseView;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox checkBox_exportCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_Data;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_check;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_language;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_changeLanguageColumn;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBox_currentLanguage;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_LanguageInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_dataInit;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_Main Ribbon_Main
        {
            get { return this.GetRibbon<Ribbon_Main>(); }
        }
    }
}
