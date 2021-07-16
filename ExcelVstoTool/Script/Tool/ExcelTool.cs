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
}