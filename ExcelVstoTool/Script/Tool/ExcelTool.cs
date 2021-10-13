using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class ExcelTool
{
    public static bool ExistSheetName(Application application , string sheetName)
    {
        foreach (Worksheet sheet in application.Worksheets)
        {
            if(sheet.Name == sheetName)
            {
                return true;
            }
        }

        return false;
    }

    public static Worksheet CreateSheet(Application application, string sheetName)
    {
        Worksheet new_wst = (Worksheet)application.Worksheets.Add();
        new_wst.Name = sheetName;

        return new_wst;
    }

    public static Worksheet GetSheet(Application application, string sheetName,bool isCreate = false)
    {
        if(ExistSheetName(application, sheetName))
        {
            return application.Worksheets[sheetName];
        }
        else
        {
            if(isCreate)
            {
                return CreateSheet(application, sheetName);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("找不到 " + sheetName);
                return null;
            }
        }
    }

    public static void ClearSheet(Worksheet sheet,bool includeFormula)
    {
        try
        {
            int rowUsed = sheet.UsedRange.Rows.Count;
            int columnUsed = sheet.UsedRange.Columns.Count;
            if (includeFormula)
            {
                //System.Windows.Forms.MessageBox.Show("includeFormula " + includeFormula);
                sheet.Range[sheet.Cells[1, 1], sheet.Cells[rowUsed, columnUsed]].Delete(XlDeleteShiftDirection.xlShiftUp);//这是删除
            }
            else
            {
                for (int i = 1; i <= rowUsed; i++)
                {
                    for (int j = 1; j <= columnUsed; j++)
                    {
                        //只删除不含公式的部分
                        if (!sheet.Cells[i, j].HasFormula)
                        {
                            sheet.Cells[i, j].Value = null;
                        }
                    }
                }
            }
        }
        catch(Exception e)
        {
            System.Windows.Forms.MessageBox.Show("ClearSheet Exception " + e.ToString());
        }
    }

    public static string Int2ColumnName(int index)
    {
        string columnName = "";

        while (index > 0)
        {
            var modulo = (index - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            index = (index - modulo) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// 从一个起点开始，返回其右边第一个空列的列序号
    /// </summary>
    public static int GetEmptyCellByCol(Worksheet sheet,int startRow,int startCol)
    {
        int row = startRow;
        int col = startCol;

        while(!string.IsNullOrEmpty( sheet.Cells[row,col].Text ))
        {
            col++;
        }

        return col;
    }

    /// <summary>
    ///  从一个起点开始，返回其下边第一个空行的行序号
    /// </summary>
    public static int GetEmptyCellByRow(Worksheet sheet, int startRow, int startCol)
    {
        int row = startRow;
        int col = startCol;

        while (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
        {
            row++;
        }

        return row;
    }

    /// <summary>
    /// 从一个起点开始，返回其右边第一个与传入值相同的列序号
    /// </summary>
    public static int FindCellByCol(Worksheet sheet, int startRow, int startCol,string value)
    {
        int row = startRow;
        int col = startCol;

        while (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
        {
            if(sheet.Cells[row, col].Text == value)
            {
                return col;
            }

            col++;
        }

        return -1;
    }

    /// <summary>
    /// 从一个起点开始，返回其下方第一个与传入值相同的行序号
    /// </summary>
    public static int FindCellByRow(Worksheet sheet, int startRow, int startCol, string value)
    {
        int row = startRow;
        int col = startCol;

        while (!string.IsNullOrEmpty(sheet.Cells[row, col].Text))
        {
            if (sheet.Cells[row, col].Text == value)
            {
                return col;
            }

            row++;
        }

        return -1;
    }

    public static string GetRangeString(Worksheet sheet, Range range)
    {
        return sheet.Name + "!" + GetRangeString(range);
    }

    public static string GetRangeString(Range range)
    {
        return Int2ColumnName(range.Column) + range.Row
            + ":"
            + Int2ColumnName(range.Column + range.Columns.Count-1) + (range.Row + +range.Rows.Count-1);
    }

    public static bool IsRepeat(Range range, string value)
    {
        for (int col = 1; col <= range.Columns.Count; col++)
        {
            for (int row = 1; row <= range.Rows.Count; row++)
            {
                string content = range[row, col].Text;
                if (content == value)
                {
                    return true;
                }
            }
        }
        return false;
    }

    public static void SaveCalcResult(Range selectRange)
    {
        Pos sPos = new Pos(selectRange);

        for (int col = 1; col <= selectRange.Columns.Count; col++)
        {
            for (int row = 1; row <= selectRange.Rows.Count; row++)
            {
                if(selectRange[row, col].HasFormula)
                {
                    selectRange[row, col] = selectRange[row, col].Text;
                }
            }

            if(col > 10000)
            {
                System.Windows.Forms.MessageBox.Show("超过1万行不再处理");
                break;
            }
        }
    }
}