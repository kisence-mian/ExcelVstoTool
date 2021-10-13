namespace ExcelVstoTool.DialogWindow
{
    partial class DataToolWindow
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.radioButton_TargetRange = new System.Windows.Forms.RadioButton();
            this.textBox_selectRange = new System.Windows.Forms.TextBox();
            this.textBox_targetRange = new System.Windows.Forms.TextBox();
            this.label_sheetName = new System.Windows.Forms.Label();
            this.label_targetRange = new System.Windows.Forms.Label();
            this.radioButton_SelectRange = new System.Windows.Forms.RadioButton();
            this.button_CompressAndExtract = new System.Windows.Forms.Button();
            this.button_Compress = new System.Windows.Forms.Button();
            this.helpProvider_tool = new System.Windows.Forms.HelpProvider();
            this.button_saveData = new System.Windows.Forms.Button();
            this.tabControl_main = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button_Neaten = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_startIndex = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox_connectChar = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.radioButton_DataAugmentOriginalID = new System.Windows.Forms.RadioButton();
            this.textBox_DataAugmentOriginalID = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.radioButton_DataAugmentOutPutPosition = new System.Windows.Forms.RadioButton();
            this.textBox_DataAugmentOutPutPosition = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.radioButton_DataAugmentNumber = new System.Windows.Forms.RadioButton();
            this.textBox_DataAugmentNumber = new System.Windows.Forms.TextBox();
            this.button_DataAugment = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.radioButton_DataAugmentRange = new System.Windows.Forms.RadioButton();
            this.textBox_DataAugmentRange = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.button_CompressAndExtractByCondition = new System.Windows.Forms.Button();
            this.radioButton_CompressCondition = new System.Windows.Forms.RadioButton();
            this.label7 = new System.Windows.Forms.Label();
            this.textBox_CompressCondition = new System.Windows.Forms.TextBox();
            this.button_CompressByCondition = new System.Windows.Forms.Button();
            this.tabControl_main.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.SuspendLayout();
            // 
            // radioButton_TargetRange
            // 
            this.radioButton_TargetRange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_TargetRange.AutoSize = true;
            this.radioButton_TargetRange.Location = new System.Drawing.Point(295, 38);
            this.radioButton_TargetRange.Name = "radioButton_TargetRange";
            this.radioButton_TargetRange.Size = new System.Drawing.Size(14, 13);
            this.radioButton_TargetRange.TabIndex = 12;
            this.radioButton_TargetRange.UseVisualStyleBackColor = true;
            this.radioButton_TargetRange.CheckedChanged += new System.EventHandler(this.radioButton_TargetRange_CheckedChanged);
            // 
            // textBox_selectRange
            // 
            this.textBox_selectRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_selectRange.Location = new System.Drawing.Point(65, 5);
            this.textBox_selectRange.Name = "textBox_selectRange";
            this.textBox_selectRange.Size = new System.Drawing.Size(215, 21);
            this.textBox_selectRange.TabIndex = 10;
            // 
            // textBox_targetRange
            // 
            this.textBox_targetRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_targetRange.Location = new System.Drawing.Point(65, 35);
            this.textBox_targetRange.Name = "textBox_targetRange";
            this.textBox_targetRange.Size = new System.Drawing.Size(215, 21);
            this.textBox_targetRange.TabIndex = 11;
            // 
            // label_sheetName
            // 
            this.label_sheetName.AutoSize = true;
            this.label_sheetName.Location = new System.Drawing.Point(5, 10);
            this.label_sheetName.Name = "label_sheetName";
            this.label_sheetName.Size = new System.Drawing.Size(53, 12);
            this.label_sheetName.TabIndex = 8;
            this.label_sheetName.Text = "选中区域";
            // 
            // label_targetRange
            // 
            this.label_targetRange.AutoSize = true;
            this.label_targetRange.Location = new System.Drawing.Point(5, 35);
            this.label_targetRange.Name = "label_targetRange";
            this.label_targetRange.Size = new System.Drawing.Size(53, 12);
            this.label_targetRange.TabIndex = 9;
            this.label_targetRange.Text = "目标区域";
            // 
            // radioButton_SelectRange
            // 
            this.radioButton_SelectRange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_SelectRange.AutoSize = true;
            this.radioButton_SelectRange.Checked = true;
            this.radioButton_SelectRange.Location = new System.Drawing.Point(295, 10);
            this.radioButton_SelectRange.Name = "radioButton_SelectRange";
            this.radioButton_SelectRange.Size = new System.Drawing.Size(14, 13);
            this.radioButton_SelectRange.TabIndex = 13;
            this.radioButton_SelectRange.TabStop = true;
            this.radioButton_SelectRange.UseVisualStyleBackColor = true;
            this.radioButton_SelectRange.CheckedChanged += new System.EventHandler(this.radioButton_SelectRange_CheckedChanged);
            // 
            // button_CompressAndExtract
            // 
            this.button_CompressAndExtract.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_CompressAndExtract.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_CompressAndExtract.Location = new System.Drawing.Point(7, 273);
            this.button_CompressAndExtract.Name = "button_CompressAndExtract";
            this.button_CompressAndExtract.Size = new System.Drawing.Size(303, 23);
            this.button_CompressAndExtract.TabIndex = 14;
            this.button_CompressAndExtract.Text = "压缩并提取";
            this.button_CompressAndExtract.UseVisualStyleBackColor = true;
            this.button_CompressAndExtract.Click += new System.EventHandler(this.button_CompressAndExtract_Click);
            // 
            // button_Compress
            // 
            this.button_Compress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Compress.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_Compress.Location = new System.Drawing.Point(7, 302);
            this.button_Compress.Name = "button_Compress";
            this.button_Compress.Size = new System.Drawing.Size(303, 23);
            this.button_Compress.TabIndex = 15;
            this.button_Compress.Text = "压缩数据";
            this.button_Compress.UseVisualStyleBackColor = true;
            this.button_Compress.Click += new System.EventHandler(this.button_Compress_Click);
            // 
            // button_saveData
            // 
            this.button_saveData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_saveData.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_saveData.Location = new System.Drawing.Point(7, 389);
            this.button_saveData.Name = "button_saveData";
            this.button_saveData.Size = new System.Drawing.Size(303, 23);
            this.button_saveData.TabIndex = 16;
            this.button_saveData.Text = "转换为计算结果";
            this.button_saveData.UseVisualStyleBackColor = true;
            this.button_saveData.Click += new System.EventHandler(this.button_saveData_Click);
            // 
            // tabControl_main
            // 
            this.tabControl_main.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl_main.Controls.Add(this.tabPage1);
            this.tabControl_main.Controls.Add(this.tabPage2);
            this.tabControl_main.Location = new System.Drawing.Point(4, 5);
            this.tabControl_main.Name = "tabControl_main";
            this.tabControl_main.SelectedIndex = 0;
            this.tabControl_main.Size = new System.Drawing.Size(323, 443);
            this.tabControl_main.TabIndex = 17;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button_Neaten);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.textBox_startIndex);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.textBox_connectChar);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.radioButton_DataAugmentOriginalID);
            this.tabPage1.Controls.Add(this.textBox_DataAugmentOriginalID);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.radioButton_DataAugmentOutPutPosition);
            this.tabPage1.Controls.Add(this.textBox_DataAugmentOutPutPosition);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.radioButton_DataAugmentNumber);
            this.tabPage1.Controls.Add(this.textBox_DataAugmentNumber);
            this.tabPage1.Controls.Add(this.button_DataAugment);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.radioButton_DataAugmentRange);
            this.tabPage1.Controls.Add(this.textBox_DataAugmentRange);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(315, 417);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "数据扩增";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button_Neaten
            // 
            this.button_Neaten.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Neaten.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_Neaten.Location = new System.Drawing.Point(6, 388);
            this.button_Neaten.Name = "button_Neaten";
            this.button_Neaten.Size = new System.Drawing.Size(303, 23);
            this.button_Neaten.TabIndex = 31;
            this.button_Neaten.Text = "整理原始ID";
            this.button_Neaten.UseVisualStyleBackColor = true;
            this.button_Neaten.Click += new System.EventHandler(this.button_Neaten_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(5, 143);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 12);
            this.label6.TabIndex = 30;
            this.label6.Text = "扩增初始序号";
            // 
            // textBox_startIndex
            // 
            this.textBox_startIndex.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_startIndex.Location = new System.Drawing.Point(88, 143);
            this.textBox_startIndex.Name = "textBox_startIndex";
            this.textBox_startIndex.Size = new System.Drawing.Size(192, 21);
            this.textBox_startIndex.TabIndex = 29;
            this.textBox_startIndex.Text = "1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(5, 116);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(65, 12);
            this.label5.TabIndex = 28;
            this.label5.Text = "扩增分隔符";
            // 
            // textBox_connectChar
            // 
            this.textBox_connectChar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_connectChar.Location = new System.Drawing.Point(88, 116);
            this.textBox_connectChar.Name = "textBox_connectChar";
            this.textBox_connectChar.Size = new System.Drawing.Size(192, 21);
            this.textBox_connectChar.TabIndex = 27;
            this.textBox_connectChar.Text = "_";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(5, 89);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 26;
            this.label4.Text = "原始ID区域";
            // 
            // radioButton_DataAugmentOriginalID
            // 
            this.radioButton_DataAugmentOriginalID.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_DataAugmentOriginalID.AutoSize = true;
            this.radioButton_DataAugmentOriginalID.Location = new System.Drawing.Point(295, 92);
            this.radioButton_DataAugmentOriginalID.Name = "radioButton_DataAugmentOriginalID";
            this.radioButton_DataAugmentOriginalID.Size = new System.Drawing.Size(14, 13);
            this.radioButton_DataAugmentOriginalID.TabIndex = 25;
            this.radioButton_DataAugmentOriginalID.UseVisualStyleBackColor = true;
            this.radioButton_DataAugmentOriginalID.CheckedChanged += new System.EventHandler(this.radioButton_DataAugmentOriginalID_CheckedChanged);
            // 
            // textBox_DataAugmentOriginalID
            // 
            this.textBox_DataAugmentOriginalID.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_DataAugmentOriginalID.Location = new System.Drawing.Point(76, 89);
            this.textBox_DataAugmentOriginalID.Name = "textBox_DataAugmentOriginalID";
            this.textBox_DataAugmentOriginalID.Size = new System.Drawing.Size(204, 21);
            this.textBox_DataAugmentOriginalID.TabIndex = 24;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 23;
            this.label3.Text = "导出区域";
            // 
            // radioButton_DataAugmentOutPutPosition
            // 
            this.radioButton_DataAugmentOutPutPosition.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_DataAugmentOutPutPosition.AutoSize = true;
            this.radioButton_DataAugmentOutPutPosition.Location = new System.Drawing.Point(295, 65);
            this.radioButton_DataAugmentOutPutPosition.Name = "radioButton_DataAugmentOutPutPosition";
            this.radioButton_DataAugmentOutPutPosition.Size = new System.Drawing.Size(14, 13);
            this.radioButton_DataAugmentOutPutPosition.TabIndex = 22;
            this.radioButton_DataAugmentOutPutPosition.UseVisualStyleBackColor = true;
            this.radioButton_DataAugmentOutPutPosition.CheckedChanged += new System.EventHandler(this.radioButton_DataAugmentOutPutPosition_CheckedChanged);
            // 
            // textBox_DataAugmentOutPutPosition
            // 
            this.textBox_DataAugmentOutPutPosition.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_DataAugmentOutPutPosition.Location = new System.Drawing.Point(76, 62);
            this.textBox_DataAugmentOutPutPosition.Name = "textBox_DataAugmentOutPutPosition";
            this.textBox_DataAugmentOutPutPosition.Size = new System.Drawing.Size(204, 21);
            this.textBox_DataAugmentOutPutPosition.TabIndex = 21;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 20;
            this.label2.Text = "扩增数目";
            // 
            // radioButton_DataAugmentNumber
            // 
            this.radioButton_DataAugmentNumber.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_DataAugmentNumber.AutoSize = true;
            this.radioButton_DataAugmentNumber.Location = new System.Drawing.Point(295, 38);
            this.radioButton_DataAugmentNumber.Name = "radioButton_DataAugmentNumber";
            this.radioButton_DataAugmentNumber.Size = new System.Drawing.Size(14, 13);
            this.radioButton_DataAugmentNumber.TabIndex = 19;
            this.radioButton_DataAugmentNumber.UseVisualStyleBackColor = true;
            this.radioButton_DataAugmentNumber.CheckedChanged += new System.EventHandler(this.radioButton_DataAugmentNumber_CheckedChanged);
            // 
            // textBox_DataAugmentNumber
            // 
            this.textBox_DataAugmentNumber.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_DataAugmentNumber.Location = new System.Drawing.Point(64, 35);
            this.textBox_DataAugmentNumber.Name = "textBox_DataAugmentNumber";
            this.textBox_DataAugmentNumber.Size = new System.Drawing.Size(216, 21);
            this.textBox_DataAugmentNumber.TabIndex = 18;
            // 
            // button_DataAugment
            // 
            this.button_DataAugment.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_DataAugment.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_DataAugment.Location = new System.Drawing.Point(6, 359);
            this.button_DataAugment.Name = "button_DataAugment";
            this.button_DataAugment.Size = new System.Drawing.Size(303, 23);
            this.button_DataAugment.TabIndex = 17;
            this.button_DataAugment.Text = "扩增数据";
            this.button_DataAugment.UseVisualStyleBackColor = true;
            this.button_DataAugment.Click += new System.EventHandler(this.button_Click_DataAugment);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 15;
            this.label1.Text = "扩增区域";
            // 
            // radioButton_DataAugmentRange
            // 
            this.radioButton_DataAugmentRange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_DataAugmentRange.AutoSize = true;
            this.radioButton_DataAugmentRange.Location = new System.Drawing.Point(295, 10);
            this.radioButton_DataAugmentRange.Name = "radioButton_DataAugmentRange";
            this.radioButton_DataAugmentRange.Size = new System.Drawing.Size(14, 13);
            this.radioButton_DataAugmentRange.TabIndex = 14;
            this.radioButton_DataAugmentRange.UseVisualStyleBackColor = true;
            this.radioButton_DataAugmentRange.CheckedChanged += new System.EventHandler(this.radioButton_DataAugmentRange_CheckedChanged);
            // 
            // textBox_DataAugmentRange
            // 
            this.textBox_DataAugmentRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_DataAugmentRange.Location = new System.Drawing.Point(63, 5);
            this.textBox_DataAugmentRange.Name = "textBox_DataAugmentRange";
            this.textBox_DataAugmentRange.Size = new System.Drawing.Size(217, 21);
            this.textBox_DataAugmentRange.TabIndex = 11;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.button_CompressAndExtractByCondition);
            this.tabPage2.Controls.Add(this.radioButton_CompressCondition);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.textBox_CompressCondition);
            this.tabPage2.Controls.Add(this.button_CompressByCondition);
            this.tabPage2.Controls.Add(this.button_CompressAndExtract);
            this.tabPage2.Controls.Add(this.button_saveData);
            this.tabPage2.Controls.Add(this.button_Compress);
            this.tabPage2.Controls.Add(this.radioButton_SelectRange);
            this.tabPage2.Controls.Add(this.textBox_selectRange);
            this.tabPage2.Controls.Add(this.radioButton_TargetRange);
            this.tabPage2.Controls.Add(this.label_targetRange);
            this.tabPage2.Controls.Add(this.label_sheetName);
            this.tabPage2.Controls.Add(this.textBox_targetRange);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(315, 417);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "数据压缩";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // button_CompressAndExtractByCondition
            // 
            this.button_CompressAndExtractByCondition.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_CompressAndExtractByCondition.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_CompressAndExtractByCondition.Location = new System.Drawing.Point(7, 331);
            this.button_CompressAndExtractByCondition.Name = "button_CompressAndExtractByCondition";
            this.button_CompressAndExtractByCondition.Size = new System.Drawing.Size(303, 23);
            this.button_CompressAndExtractByCondition.TabIndex = 21;
            this.button_CompressAndExtractByCondition.Text = "基于条件区域压缩并提取";
            this.button_CompressAndExtractByCondition.UseVisualStyleBackColor = true;
            this.button_CompressAndExtractByCondition.Click += new System.EventHandler(this.button_CompressAndExtractByCondition_Click);
            // 
            // radioButton_CompressCondition
            // 
            this.radioButton_CompressCondition.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_CompressCondition.AutoSize = true;
            this.radioButton_CompressCondition.Location = new System.Drawing.Point(295, 65);
            this.radioButton_CompressCondition.Name = "radioButton_CompressCondition";
            this.radioButton_CompressCondition.Size = new System.Drawing.Size(14, 13);
            this.radioButton_CompressCondition.TabIndex = 20;
            this.radioButton_CompressCondition.UseVisualStyleBackColor = true;
            this.radioButton_CompressCondition.CheckedChanged += new System.EventHandler(this.radioButton_CompressCondition_CheckedChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(5, 62);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(53, 12);
            this.label7.TabIndex = 18;
            this.label7.Text = "条件区域";
            // 
            // textBox_CompressCondition
            // 
            this.textBox_CompressCondition.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_CompressCondition.Location = new System.Drawing.Point(65, 62);
            this.textBox_CompressCondition.Name = "textBox_CompressCondition";
            this.textBox_CompressCondition.Size = new System.Drawing.Size(215, 21);
            this.textBox_CompressCondition.TabIndex = 19;
            // 
            // button_CompressByCondition
            // 
            this.button_CompressByCondition.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_CompressByCondition.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_CompressByCondition.Location = new System.Drawing.Point(7, 360);
            this.button_CompressByCondition.Name = "button_CompressByCondition";
            this.button_CompressByCondition.Size = new System.Drawing.Size(303, 23);
            this.button_CompressByCondition.TabIndex = 17;
            this.button_CompressByCondition.Text = "基于条件区域压缩";
            this.button_CompressByCondition.UseVisualStyleBackColor = true;
            this.button_CompressByCondition.Click += new System.EventHandler(this.button_CompressByCondition_Click);
            // 
            // DataToolWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(329, 451);
            this.Controls.Add(this.tabControl_main);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(300, 100);
            this.Name = "DataToolWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "数据工具";
            this.tabControl_main.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButton_TargetRange;
        private System.Windows.Forms.TextBox textBox_selectRange;
        private System.Windows.Forms.TextBox textBox_targetRange;
        private System.Windows.Forms.Label label_sheetName;
        private System.Windows.Forms.Label label_targetRange;
        private System.Windows.Forms.RadioButton radioButton_SelectRange;
        private System.Windows.Forms.Button button_CompressAndExtract;
        private System.Windows.Forms.Button button_Compress;
        private System.Windows.Forms.HelpProvider helpProvider_tool;
        private System.Windows.Forms.Button button_saveData;
        private System.Windows.Forms.TabControl tabControl_main;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button_DataAugment;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton radioButton_DataAugmentRange;
        private System.Windows.Forms.TextBox textBox_DataAugmentRange;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton radioButton_DataAugmentNumber;
        private System.Windows.Forms.TextBox textBox_DataAugmentNumber;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton radioButton_DataAugmentOriginalID;
        private System.Windows.Forms.TextBox textBox_DataAugmentOriginalID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RadioButton radioButton_DataAugmentOutPutPosition;
        private System.Windows.Forms.TextBox textBox_DataAugmentOutPutPosition;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_startIndex;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox_connectChar;
        private System.Windows.Forms.Button button_Neaten;
        private System.Windows.Forms.Button button_CompressByCondition;
        private System.Windows.Forms.RadioButton radioButton_CompressCondition;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textBox_CompressCondition;
        private System.Windows.Forms.Button button_CompressAndExtractByCondition;
    }
}