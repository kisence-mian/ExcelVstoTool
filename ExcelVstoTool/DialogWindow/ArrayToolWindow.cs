using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelVstoTool.DialogWindow
{
    public partial class ArrayToolWindow : Form
    {
        CheckType currentCheck = CheckType.SelectRange;

        public ArrayToolWindow()
        {
            InitializeComponent();

            helpProvider_tool.SetHelpString(button_merge, "把选中范围的数据合并成竖线分割的格式，并输出到选中区域");
            helpProvider_tool.SetHelpString(button_Expand, "把选中读取范围内的所有数组数据展开到选中区域");
            helpProvider_tool.SetHelpString(button_reverseExpand, "读取范围内的所有数据，并以引用的形式展开到指定位置(需要安装拓展公式)");
            helpProvider_tool.SetHelpString(button_saveData, "把选中读取范围内的的公式计算结果保存下来");
        }

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

        #region 外部事件接口

        public void OnSelectChange(Worksheet sheet, Range range)
        {
            string rangeString = ExcelTool.GetRangeString(sheet,range);

            if(currentCheck == CheckType.SelectRange)
            {
                textBox_selectRange.Text = rangeString;
            }
            else if(currentCheck == CheckType.TargetRange)
            {
                textBox_targetRange.Text = rangeString;
            }

            currentCheck = CheckType.No;
            radioButton_SelectRange.Checked = false;
            radioButton_TargetRange.Checked = false;
        }

        #endregion

        #region UI事件

        private void button_Expand_Click(object sender, EventArgs e)
        {
            if(!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text) 
                || !WindowUntilTool.CheckRangeFormat(textBox_targetRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取两个Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            Range targetRange = Ribbon_Main.GetRangeByRangeString(textBox_targetRange.Text);

            ExpandData(selectRange, targetRange);
        }

        private void button_merge_Click(object sender, EventArgs e)
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

            MergeData(selectRange, targetRange);
        }

        private void button_reverseExpand_Click(object sender, EventArgs e)
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

            ReverseExpandData(selectRange, targetRange);
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

        private void radioButton_SelectRange_CheckedChanged(object sender, EventArgs e)
        {
            if(radioButton_SelectRange.Checked)
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

        #endregion

        #region 展开逻辑

        void ExpandData(Range selectRange, Range targetRange)
        {
            Pos cPos = new Pos(targetRange);

            //依次读取范围内的所有数据，并展开到指定位置
            //修改原方格数据指向
            //int col = 1;
            for (int col = 1; col <= selectRange.Columns.Count; col++)
            {
                int maxArrayLength = 0;

                for (int row = 1; row <= selectRange.Rows.Count; row++)
                {
                    string content = selectRange[row, col].Text;

                    //修改原方格数据指向
                    selectRange[row, col].Formula = "=" + targetRange.Worksheet.Name + "!" + GetAimPosString(cPos, 0, row - 1);

                    string[] array = ParseTool.String2StringArray(content);

                    //修改数值
                    WriteValue(targetRange, array, cPos, row - 1);

                    //修改公式
                    targetRange.Worksheet.Cells[cPos.row + row - 1, cPos.col].Formula = GenerateArrayFormula(array, cPos, row - 1);

                    if(array.Length > maxArrayLength)
                    {
                        maxArrayLength = array.Length;
                    }
                }

                cPos.col += maxArrayLength + 1;
            }
        }



        string GetAimPosString(Pos start,int colOffset, int rowOffset)
        {
            return ExcelTool.Int2ColumnName(start.col + colOffset) + (start.row + rowOffset);
        }

        void WriteValue(Range targetRange, string[] array, Pos cPos, int rowOffset)
        {
            Worksheet sheet = targetRange.Worksheet;

            for (int i = 0; i < array.Length; i++)
            {
                sheet.Cells[cPos.row + rowOffset, cPos.col + i + 1] = array[i];
            }
        }

        string GenerateArrayFormula(string[] array,Pos cPos,int rowOffset)
        {
            if(array.Length != 0)
            {
                string result = "=";
                for (int i = 0; i < array.Length; i++)
                {
                    result += ExcelTool.Int2ColumnName(cPos.col + i + 1) + (cPos.row + rowOffset);

                    if (i != array.Length - 1)
                    {
                        result += "&\"|\"&";
                    }
                }
                return result;
            }
            else
            {
                return "";
            }
        }

        void ReverseExpandData(Range selectRange, Range targetRange)
        {
            Pos sPos = new Pos(selectRange);
            Pos cPos = new Pos(targetRange);

            string sheetName = "";

            if(selectRange.Worksheet != targetRange.Worksheet)
            {
                sheetName = selectRange.Worksheet.Name + "!";
            }

            //读取范围内的所有数据，并以引用的形式展开到指定位置
            for (int col = 1; col <= selectRange.Columns.Count; col++)
            {
                int maxArrayLength = 0;

                for (int row = 1; row <= selectRange.Rows.Count; row++)
                {
                    string content = selectRange[row, col].Text;
                    string posString = sheetName +"" + ExcelTool.Int2ColumnName(sPos.col + col -1 ) + (sPos.row + row -1);
                    string[] array = ParseTool.String2StringArray(content);

                    //修改数值
                    WriteReverseExpandValue(targetRange, posString, array, cPos, row - 1);
                }

                cPos.col += maxArrayLength + 1;
            }
        }

        void WriteReverseExpandValue(Range targetRange, string posString,string[] array, Pos cPos, int rowOffset)
        {
            Worksheet sheet = targetRange.Worksheet;

            for (int i = 0; i < array.Length; i++)
            {
                string formula = "=TextSplit(" + posString + ",\"|\"," + i + ")";
                sheet.Cells[cPos.row + rowOffset, cPos.col + i].Formula = formula;
            }
        }

        #endregion

        #region 合并逻辑

        void MergeData(Range selectRange, Range targetRange)
        {
            Worksheet targetSheet = targetRange.Worksheet;
            Worksheet setlectSheet = selectRange.Worksheet;

            string sheetName = "";

            if(targetSheet != setlectSheet)
            {
                sheetName = setlectSheet.Name + "!";
            }
            //把选中范围的 数据合并成竖线分割的格式，并输出到选中区域

            for (int row = 1; row <= selectRange.Rows.Count; row++)
            {
                Pos sPos = new Pos(selectRange);
                string formuale = "=";
                bool isAdd = false;
                for (int col = 1; col <= selectRange.Columns.Count; col++)
                {
                    string content = setlectSheet.Cells[selectRange.Row + row - 1, selectRange.Column + col - 1].Text;
                    if (!string.IsNullOrEmpty(content))
                    {
                        if (isAdd)
                        {
                            formuale += "&\"|\"&";
                        }
                        else
                        {
                            isAdd = true;
                        }

                        formuale += sheetName + ExcelTool.Int2ColumnName(selectRange.Column + col - 1) + (selectRange.Row + row - 1);
                    }
                }

                //写入数据
                targetSheet.Cells[targetRange.Row + row -1, targetRange.Column].Formula = formuale;
            }
        }
        #endregion


    }
}
