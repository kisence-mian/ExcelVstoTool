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
            helpProvider_tool.SetHelpString(button_Expand, "把选中读取范围内的所有数组数据，并展开到选中区域");
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
            string rangeString = GetRangeString(sheet,range);

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
            if(!CheckRangeFormat(textBox_selectRange.Text) 
                || !CheckRangeFormat(textBox_targetRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取两个Range
            Range selectRange = GetRangeByRangeString(textBox_selectRange.Text);
            Range targetRange = GetRangeByRangeString(textBox_targetRange.Text);

            ExpandData(selectRange, targetRange);
        }

        private void button_merge_Click(object sender, EventArgs e)
        {
            if (!CheckRangeFormat(textBox_selectRange.Text)
                || !CheckRangeFormat(textBox_targetRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取两个Range
            Range selectRange = GetRangeByRangeString(textBox_selectRange.Text);
            Range targetRange = GetRangeByRangeString(textBox_targetRange.Text);

            MergeData(selectRange, targetRange);
        }

        bool CheckRangeFormat(string content)
        {
            return Regex.IsMatch(content, "^([\\s\\S]*)![A-Z]+[0-9]+:[A-Z]+[0-9]+$");
        }


        private void radioButton_SelectRange_CheckedChanged(object sender, EventArgs e)
        {
            currentCheck = CheckType.SelectRange;
            textBox_selectRange.Text =  GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
        }

        private void radioButton_TargetRange_CheckedChanged(object sender, EventArgs e)
        {
            currentCheck = CheckType.TargetRange;
            textBox_targetRange.Text = GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
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
            string result = "=";
            for (int i = 0; i < array.Length; i++)
            {
                result += ExcelTool.Int2ColumnName(cPos.col + i+1) + (cPos.row  + rowOffset);

                if(i != array.Length -1)
                {
                    result += "&\"|\"&";
                }
            }

            return result;
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
                for (int col = 1; col <= selectRange.Columns.Count; col++)
                {
                    formuale += sheetName + ExcelTool.Int2ColumnName(selectRange.Column + col - 1) + (selectRange.Row + row-1);

                    if(col != selectRange.Columns.Count)
                    {
                        formuale += "&\"|\"&";
                    }
                }

                //写入数据
                targetSheet.Cells[targetRange.Row + row -1, targetRange.Column].Formula = formuale;
            }
        }


        #endregion

        #region 工具方法

        string GetRangeString(Worksheet sheet, Range range)
        {
            return sheet.Name + "!" + ExcelTool.GetRangeString(range);
        }

        Range GetRangeByRangeString(string rangeString)
        {
            string SheetName = rangeString.Split('!')[0];
            string range = rangeString.Split('!')[1];

            Worksheet worksheet = Ribbon_Main.GetSheet(SheetName, false);

            return worksheet.Range[range];
        }

        #endregion


    }

    enum CheckType
    {
        SelectRange,
        TargetRange,
        No,
    }

    struct Pos
    {
        public int col;
        public int row;
        
        public Pos(Range range)
        {
            col = range.Column;
            row = range.Row;
        }
    }
}
