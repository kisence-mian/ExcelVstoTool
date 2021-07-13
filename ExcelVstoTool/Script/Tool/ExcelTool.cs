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
}