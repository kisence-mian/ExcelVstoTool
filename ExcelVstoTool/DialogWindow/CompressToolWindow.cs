using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelVstoTool.DialogWindow
{
    public partial class CompressToolWindow : Form
    {
        CheckType currentCheck = CheckType.SelectRange;

        public CompressToolWindow()
        {
            InitializeComponent();

            helpProvider_tool.SetHelpString(button_Compress, "从上而下遍历选中区域，将相同的项改为对第一个项的引用");
            helpProvider_tool.SetHelpString(button_CompressAndExtract, "从上而下遍历选中区域，将相同的项改为对第一个项的引用,并将结果输出到目标区域");
            helpProvider_tool.SetHelpString(button_saveData, "把选中读取范围内的的公式计算结果保存下来");
        }

        #region 生命周期

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Ribbon_Main.OnSelectChange += OnSelectChange;
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            base.OnHandleDestroyed(e);

            Ribbon_Main.OnSelectChange -= OnSelectChange;
        }

        #endregion

        #region 外部事件接口

        public void OnSelectChange(Worksheet sheet, Range range)
        {
            string rangeString = ExcelTool.GetRangeString(sheet, range);

            if (currentCheck == CheckType.SelectRange)
            {
                textBox_selectRange.Text = rangeString;
            }
            else if (currentCheck == CheckType.TargetRange)
            {
                textBox_targetRange.Text = rangeString;
            }

            currentCheck = CheckType.No;
            radioButton_SelectRange.Checked = false;
            radioButton_TargetRange.Checked = false;
        }

        #endregion

        #region UI事件

        private void button_CompressAndExtract_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_targetRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取两个Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            Range targetRange = Ribbon_Main.GetRangeByRangeString(textBox_targetRange.Text);

            CompressDataAndExtract(selectRange, targetRange);
        }

        private void button_Compress_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            CompressData(selectRange);
        }

        private void radioButton_SelectRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_SelectRange.Checked)
            {
                currentCheck = CheckType.SelectRange;
                textBox_selectRange.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void radioButton_TargetRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_TargetRange.Checked)
            {
                currentCheck = CheckType.TargetRange;
                textBox_targetRange.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void button_saveData_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            ExcelTool.SaveCalcResult(selectRange);
        }

        #endregion

        #region 压缩与提取

        void CompressData(Range selectRange)
        {
            //从上而下遍历选中区域，将相同的项改为对第一个项的引用
            Pos cPos = new Pos(selectRange);
            for (int col = 1; col <= selectRange.Columns.Count; col++)
            {
                int rRow = 1;
                string content = null;
                for (int row = 1; row <= selectRange.Rows.Count; row++)
                {
                    if(selectRange[row,col].Text == content)
                    {
                        selectRange[row, col].Formula = "=" + ExcelTool.Int2ColumnName(cPos.col + col - 1) + (cPos.row + rRow - 1);
                    }
                    else
                    {
                        content = selectRange[row, col].Text;
                        rRow = row;
                    }
                }
            }
        }

        void CompressDataAndExtract(Range selectRange,Range targetRange)
        {
            Pos cPos = new Pos(selectRange);
            Pos tPos = new Pos(targetRange); //指向写入的位置

            string sheetName = "";
            if(targetRange.Worksheet != selectRange.Worksheet)
            {
                sheetName = targetRange.Worksheet.Name + "!";
            }

            Worksheet targetSheet = targetRange.Worksheet;
            //从上而下遍历选中区域，将相同的项改为对第一个项的引用,并将不重复的结果输出到选中区域

            for (int col = 1; col <= selectRange.Columns.Count; col++)
            {
                int rRow = 1;
                string content = null;
                for (int row = 1; row <= selectRange.Rows.Count; row++)
                {
                    if (selectRange[row, col].Text == content)
                    {
                        selectRange[row, col].Formula = "=" + ExcelTool.Int2ColumnName(cPos.col + col - 1) + (cPos.row + rRow - 1);
                    }
                    else
                    {
                        content = selectRange[row, col].Text;
                        rRow = row;

                        //输出到目标区域,并将引用指向过去
                        targetSheet.Cells[tPos.row,tPos.col] = selectRange[row, col].Text;
                        selectRange[row, col].Formula = "=" + sheetName +  ExcelTool.Int2ColumnName(tPos.col) + (tPos.row);
                        tPos.row += 1;
                    }
                }

                tPos.col += 1;
                tPos.row = targetRange.Row;
            }
        }

        #endregion


    }
}
