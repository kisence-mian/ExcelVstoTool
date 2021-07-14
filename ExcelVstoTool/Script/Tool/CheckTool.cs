using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class CheckTool
{
    public static bool CheckSheet(Worksheet workSheet,DataConfig config)
    {
        bool result = true;

        try
        {
            //表头校验
            result &= CheckTitle(workSheet, config);

            //格式校验
            DataTable data = DataTool.Excel2Table(workSheet, config);

            //外部校验

        }
        catch (Exception e)
        {
            System.Windows.Forms.MessageBox.Show("校验出错 ->" + config.m_sheetName + " \n" + e.Message);
            return false;
        }
        return result;
    }

    static bool CheckTitle(Worksheet workSheet, DataConfig config)
    {
        bool result = true;

        if (string.IsNullOrEmpty( workSheet.Range["A1"].Value ))
        {
            throw new Exception("没有找到 表头 ");
        }

        if (workSheet.Range["A2"].Value != DataTable.c_fieldTypeTableTitle)
        {
            throw new Exception("没有找到 类型 声明行 " + DataTable.c_fieldTypeTableTitle);
        }

        if (workSheet.Range["A3"].Value != DataTable.c_noteTableTitle)
        {
            throw new Exception("没有找到 描述 声明行 " + DataTable.c_noteTableTitle);
        }

        if (workSheet.Range["A4"].Value != DataTable.c_defaultValueTableTitle)
        {
            throw new Exception("没有找到 默认值 声明行 " + DataTable.c_defaultValueTableTitle);
        }

        return result;
    }

    static bool CheckColumn(Worksheet workSheet,int col, DataConfig config)
    {
        bool result = true;

        //读取类型


        return result;
    }
}
