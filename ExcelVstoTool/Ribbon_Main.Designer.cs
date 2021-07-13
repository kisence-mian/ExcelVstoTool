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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button_initData = this.Factory.CreateRibbonButton();
            this.button_ToExcel = this.Factory.CreateRibbonButton();
            this.button_toTxt = this.Factory.CreateRibbonButton();
            this.tab_main.SuspendLayout();
            this.group_main.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab_main
            // 
            this.tab_main.Groups.Add(this.group_main);
            this.tab_main.Groups.Add(this.group1);
            this.tab_main.Label = "拓展工具";
            this.tab_main.Name = "tab_main";
            this.tab_main.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            // 
            // group_main
            // 
            this.group_main.Items.Add(this.button_initData);
            this.group_main.Label = "设置";
            this.group_main.Name = "group_main";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button_ToExcel);
            this.group1.Items.Add(this.button_toTxt);
            this.group1.Label = "操作";
            this.group1.Name = "group1";
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_main;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group_main;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_toTxt;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_ToExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button_initData;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon_Main Ribbon_Main
        {
            get { return this.GetRibbon<Ribbon_Main>(); }
        }
    }
}
