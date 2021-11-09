namespace ExcelVstoTool.DialogWindow
{
    partial class FileToolWindow
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
            this.button_batchExcel2Text = new System.Windows.Forms.Button();
            this.button_batchText2Excel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_batchExcel2Text
            // 
            this.button_batchExcel2Text.Location = new System.Drawing.Point(12, 12);
            this.button_batchExcel2Text.Name = "button_batchExcel2Text";
            this.button_batchExcel2Text.Size = new System.Drawing.Size(240, 23);
            this.button_batchExcel2Text.TabIndex = 0;
            this.button_batchExcel2Text.Text = "批量转换 Excel 到 txt";
            this.button_batchExcel2Text.UseVisualStyleBackColor = true;
            this.button_batchExcel2Text.Click += new System.EventHandler(this.button_batchExcel2Text_Click);
            // 
            // button_batchText2Excel
            // 
            this.button_batchText2Excel.Location = new System.Drawing.Point(12, 41);
            this.button_batchText2Excel.Name = "button_batchText2Excel";
            this.button_batchText2Excel.Size = new System.Drawing.Size(240, 23);
            this.button_batchText2Excel.TabIndex = 1;
            this.button_batchText2Excel.Text = "批量导入 txt 到 Excel";
            this.button_batchText2Excel.UseVisualStyleBackColor = true;
            this.button_batchText2Excel.Click += new System.EventHandler(this.button_batchText2Excel_Click);
            // 
            // FileToolWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(264, 450);
            this.Controls.Add(this.button_batchText2Excel);
            this.Controls.Add(this.button_batchExcel2Text);
            this.HelpButton = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FileToolWindow";
            this.ShowIcon = false;
            this.Text = "文件操作";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button_batchExcel2Text;
        private System.Windows.Forms.Button button_batchText2Excel;
    }
}