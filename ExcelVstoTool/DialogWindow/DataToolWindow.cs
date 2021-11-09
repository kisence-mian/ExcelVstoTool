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
    public partial class DataToolWindow : Form
    {
        public DataToolWindow()
        {
            InitializeComponent();

            helpProvider_tool.SetHelpString(button_CompressAndExtract, "从上而下遍历选中区域，将相同的项改为对第一个项的引用,并将结果输出到目标区域");
            helpProvider_tool.SetHelpString(button_Compress, "从上而下遍历选中区域，将相同的项改为对第一个项的引用");
            helpProvider_tool.SetHelpString(button_CompressAndExtractByCondition, "从上而下遍历选中区域，将对应位置条件区域中相同的项改为对第一个项的引用,并将结果输出到目标区域");
            helpProvider_tool.SetHelpString(button_CompressByCondition, "从上而下遍历选中区域，将对应位置条件区域中相同的项改为对第一个项的引用");
            helpProvider_tool.SetHelpString(button_saveData, "把选中读取范围内的的公式计算结果保存下来");
            helpProvider_tool.SetHelpString(button_RemoveRepeatExtract, "把选中读取范围内的值无重复地提取到目标位置");

            helpProvider_tool.SetHelpString(button_DataAugment, "把扩增区域中的数据增加连接符和序号ID，每个重复n次(读取扩增数目区域)，输出到导出位置");
            helpProvider_tool.SetHelpString(button_Neaten, "把扩增区域中的数据整理到原始ID区域，并自动与旧数据进行去重");
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

            //数据扩增

            if (radioButton_DataAugmentRange.Checked)
            {
                textBox_DataAugmentRange.Text = rangeString;
                radioButton_DataAugmentRange.Checked = false;
            }

            if (radioButton_DataAugmentNumber.Checked)
            {
                textBox_DataAugmentNumber.Text = rangeString;
                radioButton_DataAugmentNumber.Checked = false;
            }

            if (radioButton_DataAugmentOutPutPosition.Checked)
            {
                textBox_DataAugmentOutPutPosition.Text = rangeString;
                radioButton_DataAugmentOutPutPosition.Checked = false;
            }

            if (radioButton_DataAugmentOriginalID.Checked)
            {
                textBox_DataAugmentOriginalID.Text = rangeString;
                radioButton_DataAugmentOriginalID.Checked = false;
            }

            //数据压缩

            if (radioButton_SelectRange.Checked)
            {
                textBox_selectRange.Text = rangeString;
                radioButton_SelectRange.Checked = false;
            }

            if (radioButton_TargetRange.Checked)
            {
                textBox_targetRange.Text = rangeString;
                radioButton_TargetRange.Checked = false;
            }

            if (radioButton_CompressCondition.Checked)
            {
                textBox_CompressCondition.Text = rangeString;
                radioButton_CompressCondition.Checked = false;
            }


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

        private void button_CompressAndExtractByCondition_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_targetRange.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_CompressCondition.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            Range targetRange = Ribbon_Main.GetRangeByRangeString(textBox_targetRange.Text);
            Range conditionRange = Ribbon_Main.GetRangeByRangeString(textBox_CompressCondition.Text);
            CompressDataByCondition(selectRange, targetRange, conditionRange);
        }

        private void button_CompressByCondition_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_CompressCondition.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            Range conditionRange = Ribbon_Main.GetRangeByRangeString(textBox_CompressCondition.Text);
            CompressDataByCondition(selectRange, conditionRange);
        }

        private void button_RemoveRepeatExtract_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_selectRange.Text)
            || !WindowUntilTool.CheckRangeFormat(textBox_targetRange.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取Range
            Range selectRange = Ribbon_Main.GetRangeByRangeString(textBox_selectRange.Text);
            Range targetRange = Ribbon_Main.GetRangeByRangeString(textBox_targetRange.Text);
            RemoveRepeatExtract(selectRange, targetRange);
        }

        private void radioButton_SelectRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_SelectRange.Checked)
            {
                textBox_selectRange.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void radioButton_TargetRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_TargetRange.Checked)
            {
                textBox_targetRange.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void radioButton_CompressCondition_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_CompressCondition.Checked)
            {
                textBox_CompressCondition.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
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

        private void radioButton_DataAugmentRange_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_DataAugmentRange.Checked)
            {
                textBox_DataAugmentRange.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void radioButton_DataAugmentNumber_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_DataAugmentNumber.Checked)
            {
                textBox_DataAugmentNumber.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void radioButton_DataAugmentOutPutPosition_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_DataAugmentOutPutPosition.Checked)
            {
                textBox_DataAugmentOutPutPosition.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void radioButton_DataAugmentOriginalID_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_DataAugmentOriginalID.Checked)
            {
                textBox_DataAugmentOriginalID.Text = ExcelTool.GetRangeString(Ribbon_Main.GetActiveSheet(), Ribbon_Main.GetCurrentSelectRange());
            }
        }

        private void button_Click_DataAugment(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_DataAugmentRange.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_DataAugmentNumber.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_DataAugmentOutPutPosition.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_DataAugmentOriginalID.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            //获取四个Range
            Range dataAugmentRange = Ribbon_Main.GetRangeByRangeString(textBox_DataAugmentRange.Text);
            Range dataAugmentNumber = Ribbon_Main.GetRangeByRangeString(textBox_DataAugmentNumber.Text);
            Range dataAugmentOutPutPosition = Ribbon_Main.GetRangeByRangeString(textBox_DataAugmentOutPutPosition.Text);
            Range dataAugmentOriginalID = Ribbon_Main.GetRangeByRangeString(textBox_DataAugmentOriginalID.Text);

            string connectChar = textBox_connectChar.Text;
            int startIndex = int.Parse(textBox_startIndex.Text);

            Ribbon_Main.PerformanceSwitch(true);

            DataAugment(dataAugmentRange, 
                dataAugmentNumber,
                dataAugmentOutPutPosition, 
                dataAugmentOriginalID, 
                connectChar,
                startIndex);

            Ribbon_Main.PerformanceSwitch(false);
        }

        private void button_Neaten_Click(object sender, EventArgs e)
        {
            if (!WindowUntilTool.CheckRangeFormat(textBox_DataAugmentRange.Text)
                || !WindowUntilTool.CheckRangeFormat(textBox_DataAugmentOriginalID.Text))
            {
                MessageBox.Show("不正确的范围格式");
                return;
            }

            Range dataAugmentRange = Ribbon_Main.GetRangeByRangeString(textBox_DataAugmentRange.Text);
            Range dataAugmentOriginalID = Ribbon_Main.GetRangeByRangeString(textBox_DataAugmentOriginalID.Text);

            DataNeaten(dataAugmentRange, dataAugmentOriginalID);
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
        void CompressDataByCondition(Range selectRange, Range targetRange, Range conditionRange)
        {
            Pos cPos = new Pos(selectRange);
            Pos tPos = new Pos(targetRange); //指向写入的位置

            string sheetName = "";
            if (targetRange.Worksheet != selectRange.Worksheet)
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
                    int colTemp = col;

                    if (conditionRange.Columns.Count < colTemp)
                    {
                        colTemp = conditionRange.Columns.Count;
                    }

                    if (conditionRange[row, colTemp].Text == content)
                    {
                        selectRange[row, col].Formula = "=" + ExcelTool.Int2ColumnName(cPos.col + col - 1) + (cPos.row + rRow - 1);
                    }
                    else
                    {
                        content = conditionRange[row, colTemp].Text;
                        rRow = row;

                        //输出到目标区域,并将引用指向过去
                        targetSheet.Cells[tPos.row, tPos.col] = selectRange[row, col].Text;
                        selectRange[row, col].Formula = "=" + sheetName + ExcelTool.Int2ColumnName(tPos.col) + (tPos.row);
                        tPos.row += 1;
                    }
                }

                tPos.col += 1;
                tPos.row = targetRange.Row;
            }
        }

        void CompressDataByCondition(Range selectRange,Range conditionRange)
        {
            //从上而下遍历选中区域，将相同的项改为对第一个项的引用
            Pos cPos = new Pos(selectRange);
            for (int col = 1; col <= selectRange.Columns.Count; col++)
            {
                int rRow = 1;
                string content = null;
                for (int row = 1; row <= selectRange.Rows.Count; row++)
                {
                    int colTemp = col;

                    if(conditionRange.Columns.Count < colTemp)
                    {
                        colTemp = conditionRange.Columns.Count;
                    }

                    if (conditionRange[row, colTemp].Text == content)
                    {
                        selectRange[row, col].Formula = "=" + ExcelTool.Int2ColumnName(cPos.col + col - 1) + (cPos.row + rRow - 1);
                    }
                    else
                    {
                        content = conditionRange[row, colTemp].Text;
                        rRow = row;
                    }
                }
            }
        }

        void RemoveRepeatExtract(Range selectRange, Range targetRange)
        {
            Pos pos = new Pos(targetRange);
            Worksheet worksheet = targetRange.Worksheet;
            List<string> contents = new List<string>();

            for (int col = 1; col <= selectRange.Columns.Count; col++)
            {
                for (int row = 1; row <= selectRange.Rows.Count; row++)
                {
                    string content = selectRange[row, col].Text;

                    //不重复地进行收集
                    if(!contents.Contains(content))
                    {
                        contents.Add(content);
                    }
                }
            }

            for (int i = 0; i < contents.Count; i++)
            {
                worksheet.Cells[pos.row + i,pos.col] = contents[i];
            }
        }

        #endregion

        #region 数据扩增

        void DataAugment(Range dataAugmentRange, Range numberRange,Range outPutPosition, Range originKeyRange,string connectChar,int startIndex)
        {
            DateTime now = System.DateTime.Now;
            string info = "扩增完毕";

            //先收集旧数据
            DataNeatenList oldData = new DataNeatenList(outPutPosition);

            Pos outPos = new Pos(outPutPosition);
            Pos outKeyPos = new Pos(originKeyRange);

            Worksheet outPosSheet = outPutPosition.Worksheet;
            Worksheet outKeyPosSheet = originKeyRange.Worksheet;

            for (int col = 1; col <= dataAugmentRange.Columns.Count; col++)
            {
                for (int row = 1; row <= dataAugmentRange.Rows.Count; row++)
                {
                    string key = dataAugmentRange[row, col].Text;
                    if (!string.IsNullOrEmpty(key))
                    {
                        //读取波数
                        int count = 10;
                        int colTemp = col;
                        //自动适应读取哪一列
                        if(numberRange.Columns.Count < colTemp)
                        {
                            colTemp = numberRange.Columns.Count;
                        }
                        count = int.Parse( numberRange.Cells[row, colTemp].Text);

                        for (int i = 0; i < count; i++)
                        {
                            string content = key + connectChar + (i + startIndex);

                            if(!oldData.IsRepeat(content))
                            {
                                outPosSheet.Range[outPos.row + ":" + outPos.row].Insert(XlDirection.xlDown);
                                outPosSheet.Cells[outPos.row, outPos.col] = content;

                                if(outPosSheet != outKeyPosSheet)
                                {
                                    outKeyPosSheet.Range[outPos.row + ":" + outPos.row].Insert(XlDirection.xlDown);
                                }

                                outKeyPosSheet.Cells[outKeyPos.row, outKeyPos.col] = key;

                                oldData.AddRow(outPos.row);
                            }

                            outPos.row++;
                            outKeyPos.row++;
                        }
                    }
                }
            }

            info += "\n用时：" + (DateTime.Now - now).TotalSeconds + "s";

            int noUseCount = 0;
            //收集没再使用的行并使其变色
            for (int i = 0; i < oldData.Count; i++)
            {
                if (!oldData[i].isUse)
                {
                    try
                    {                     //变色
                        outPosSheet.Cells[oldData[i].row, oldData[i].col].AddComment("不曾使用的Key,请检查表格");
                    }
                    catch {}

                    noUseCount++;
                }
            }

            if(noUseCount > 0)
            {
                info += "\n存在" + noUseCount + "条不匹配的Key";
            }

            MessageBox.Show(info);
        }

        void DataNeaten(Range dataAugmentRange, Range originKeyRange)
        {
            //先收集旧数据
            DataNeatenList oldData = new DataNeatenList(originKeyRange);

            Pos outPos = new Pos(originKeyRange);

            Worksheet outPosSheet = originKeyRange.Worksheet;

            for (int col = 1; col <= dataAugmentRange.Columns.Count; col++)
            {
                for (int row = 1; row <= dataAugmentRange.Rows.Count; row++)
                {
                    string key = dataAugmentRange[row, col].Text;

                    if (!string.IsNullOrEmpty(key) )
                    {
                        if(!oldData.IsRepeat( key))
                        {
                            outPosSheet.Range[outPos.row + ":" + outPos.row].Insert(XlDirection.xlDown);
                            outPosSheet.Cells[outPos.row, outPos.col] = key;

                            oldData.AddRow(outPos.row);
                        }

                        outPos.row++;
                    }
                }
            }

            //收集没再使用的行并使其变色
            for (int i = 0; i < oldData.Count; i++)
            {
                if(!oldData[i].isUse)
                {
                    try
                    {
                        //变色
                        outPosSheet.Cells[oldData[i].row, oldData[i].col].AddComment("不曾使用的Key,请检查表格");
                    }
                    catch { }
                }
            }
        }

        class NeatonData
        {
            public string content;
            public int row;
            public int col;
            public bool isUse;

            public NeatonData(string content, int row,int col)
            {
                this.content = content;
                this.row = row;
                this.col = col;
                isUse = false;
            }
        }

        class DataNeatenList : List<NeatonData>
        {
            public DataNeatenList(Range range)
            {
                Pos pos = new Pos(range);
                Worksheet worksheet = range.Worksheet;

                while (!string.IsNullOrEmpty(worksheet.Cells[pos.row,1].Text))
                {
                    NeatonData data = new NeatonData(worksheet.Cells[pos.row, 1].Text, pos.row,pos.col);
                    Add(data);

                    pos.row++;
                }
            }

            public bool IsRepeat(string value)
            {
                for (int i = 0; i < Count; i++)
                {
                    NeatonData data = this[i];
                    if (data.content == value)
                    {
                        data.isUse = true;
                        return true;
                    }
                }

                return false;
            }

            /// <summary>
            /// 插入一行数据
            /// </summary>
            /// <param name="index"></param>
            public void AddRow(int index)
            {
                for (int i = index; i < Count; i++)
                {
                    NeatonData data = this[i];
                    data.row ++;
                }
            }
        }





        #endregion


    }
}
