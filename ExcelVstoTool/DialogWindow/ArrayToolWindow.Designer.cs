namespace ExcelVstoTool.DialogWindow
{
    partial class ArrayToolWindow
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
            this.label_sheetName = new System.Windows.Forms.Label();
            this.button_Expand = new System.Windows.Forms.Button();
            this.radioButton_TargetRange = new System.Windows.Forms.RadioButton();
            this.radioButton_SelectRange = new System.Windows.Forms.RadioButton();
            this.textBox_targetRange = new System.Windows.Forms.TextBox();
            this.textBox_selectRange = new System.Windows.Forms.TextBox();
            this.label_targetRange = new System.Windows.Forms.Label();
            this.button_merge = new System.Windows.Forms.Button();
            this.helpProvider_tool = new System.Windows.Forms.HelpProvider();
            this.button_saveData = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label_sheetName
            // 
            this.label_sheetName.AutoSize = true;
            this.label_sheetName.Location = new System.Drawing.Point(12, 15);
            this.label_sheetName.Name = "label_sheetName";
            this.label_sheetName.Size = new System.Drawing.Size(53, 12);
            this.label_sheetName.TabIndex = 1;
            this.label_sheetName.Text = "选中区域";
            // 
            // button_Expand
            // 
            this.button_Expand.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_Expand.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_Expand.Location = new System.Drawing.Point(14, 358);
            this.button_Expand.Name = "button_Expand";
            this.button_Expand.Size = new System.Drawing.Size(312, 23);
            this.button_Expand.TabIndex = 2;
            this.button_Expand.Text = "展开数组";
            this.button_Expand.UseVisualStyleBackColor = true;
            this.button_Expand.Click += new System.EventHandler(this.button_Expand_Click);
            // 
            // radioButton_TargetRange
            // 
            this.radioButton_TargetRange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_TargetRange.AutoSize = true;
            this.radioButton_TargetRange.Location = new System.Drawing.Point(312, 46);
            this.radioButton_TargetRange.Name = "radioButton_TargetRange";
            this.radioButton_TargetRange.Size = new System.Drawing.Size(14, 13);
            this.radioButton_TargetRange.TabIndex = 7;
            this.radioButton_TargetRange.UseVisualStyleBackColor = true;
            this.radioButton_TargetRange.CheckedChanged += new System.EventHandler(this.radioButton_TargetRange_CheckedChanged);
            // 
            // radioButton_SelectRange
            // 
            this.radioButton_SelectRange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.radioButton_SelectRange.AutoSize = true;
            this.radioButton_SelectRange.Checked = true;
            this.radioButton_SelectRange.Location = new System.Drawing.Point(312, 15);
            this.radioButton_SelectRange.Name = "radioButton_SelectRange";
            this.radioButton_SelectRange.Size = new System.Drawing.Size(14, 13);
            this.radioButton_SelectRange.TabIndex = 6;
            this.radioButton_SelectRange.TabStop = true;
            this.radioButton_SelectRange.UseVisualStyleBackColor = true;
            this.radioButton_SelectRange.CheckedChanged += new System.EventHandler(this.radioButton_SelectRange_CheckedChanged);
            // 
            // textBox_targetRange
            // 
            this.textBox_targetRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_targetRange.Location = new System.Drawing.Point(71, 43);
            this.textBox_targetRange.Name = "textBox_targetRange";
            this.textBox_targetRange.Size = new System.Drawing.Size(235, 21);
            this.textBox_targetRange.TabIndex = 5;
            // 
            // textBox_selectRange
            // 
            this.textBox_selectRange.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_selectRange.Location = new System.Drawing.Point(71, 12);
            this.textBox_selectRange.Name = "textBox_selectRange";
            this.textBox_selectRange.Size = new System.Drawing.Size(235, 21);
            this.textBox_selectRange.TabIndex = 4;
            // 
            // label_targetRange
            // 
            this.label_targetRange.AutoSize = true;
            this.label_targetRange.Location = new System.Drawing.Point(12, 47);
            this.label_targetRange.Name = "label_targetRange";
            this.label_targetRange.Size = new System.Drawing.Size(53, 12);
            this.label_targetRange.TabIndex = 3;
            this.label_targetRange.Text = "目标区域";
            // 
            // button_merge
            // 
            this.button_merge.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_merge.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_merge.Location = new System.Drawing.Point(13, 387);
            this.button_merge.Name = "button_merge";
            this.button_merge.Size = new System.Drawing.Size(312, 23);
            this.button_merge.TabIndex = 8;
            this.button_merge.Text = "合并数组";
            this.button_merge.UseVisualStyleBackColor = true;
            this.button_merge.Click += new System.EventHandler(this.button_merge_Click);
            // 
            // button_saveData
            // 
            this.button_saveData.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_saveData.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.button_saveData.Location = new System.Drawing.Point(12, 416);
            this.button_saveData.Name = "button_saveData";
            this.button_saveData.Size = new System.Drawing.Size(312, 23);
            this.button_saveData.TabIndex = 9;
            this.button_saveData.Text = "转换为计算结果";
            this.button_saveData.UseVisualStyleBackColor = true;
            this.button_saveData.Click += new System.EventHandler(this.button_saveData_Click);
            // 
            // ArrayToolWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(338, 450);
            this.Controls.Add(this.button_saveData);
            this.Controls.Add(this.button_merge);
            this.Controls.Add(this.button_Expand);
            this.Controls.Add(this.radioButton_TargetRange);
            this.Controls.Add(this.radioButton_SelectRange);
            this.Controls.Add(this.textBox_selectRange);
            this.Controls.Add(this.textBox_targetRange);
            this.Controls.Add(this.label_sheetName);
            this.Controls.Add(this.label_targetRange);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ArrayToolWindow";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "数组工具";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label_sheetName;
        private System.Windows.Forms.Button button_Expand;
        private System.Windows.Forms.TextBox textBox_targetRange;
        private System.Windows.Forms.TextBox textBox_selectRange;
        private System.Windows.Forms.Label label_targetRange;
        private System.Windows.Forms.RadioButton radioButton_TargetRange;
        private System.Windows.Forms.RadioButton radioButton_SelectRange;
        private System.Windows.Forms.Button button_merge;
        private System.Windows.Forms.HelpProvider helpProvider_tool;
        private System.Windows.Forms.Button button_saveData;
    }
}