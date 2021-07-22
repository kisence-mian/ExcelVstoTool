using ExcelVstoTool;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

public class DataConfig
{
    int m_index; 

    public string m_sheetName;
    public string m_txtName;

    //public string m_txtPath;

    //导入时覆盖Excel公式
    public bool m_coverFormula;

    public string GetTextPath()
    {
        return PathCovert(m_txtName);
    }

    public DataConfig()
    {

    }

    public DataConfig(Worksheet configSheet, int index)
    {
        m_index = index;

        string key             = "A" + index;
        string key_txtPath     = "B" + index;
        string key_saveFormula = "C" + index;

        m_sheetName = configSheet.Range[key].Text;
        m_txtName = configSheet.Range[key_txtPath].Text;

        m_coverFormula = (bool)ParseTool.GetBool( configSheet.Range[key_saveFormula].Text);
    }

    public static void ConfigInit(Worksheet configSheet)
    {
        configSheet.Range["A1"].Value = "页签名称";
        configSheet.Range["B1"].Value = "文件名";
        configSheet.Range["C1"].Value = "覆盖公式";

        configSheet.Range["E1"].Value = "枚举设置";
    }

    public static void AddSheetConfig(Worksheet configSheet,DataConfig config)
    {
        //找到可以写入的位置
        int index = 2;
        while(!string.IsNullOrEmpty(configSheet.Range["A" + index].Text))
        {
            index++;
        }

        config.m_index = index;

        //进行写入
        configSheet.Range["A" + index].Value = config.m_sheetName;
        configSheet.Range["B" + index].Value = config.m_txtName;
        configSheet.Range["C" + index].Value = config.m_coverFormula;
    }

    public void Delete(Worksheet configSheet)
    {
        configSheet.Range["A" + m_index + ":C" + m_index].Delete(Microsoft.Office.Interop.Excel.XlDirection.xlUp);
    }

    /// <summary>
    /// 判断一个页面是否是工作页面
    /// </summary>
    public static bool IsWorkSheet(Worksheet configSheet,string sheetName)
    {
        int index = 2;
        while (!string.IsNullOrEmpty(configSheet.Range["A" + index].Text))
        {
            if(configSheet.Range["A" + index].Text == sheetName)
            {
                return true;
            }

            index++;
        }

        return false;
    }

    public static int GetWorkIndex(Worksheet configSheet, string sheetName)
    {
        int index = 2;
        while (!string.IsNullOrEmpty(configSheet.Range["A" + index].Text))
        {
            if (configSheet.Range["A" + index].Text == sheetName)
            {
                return index;
            }

            index++;
        }

        return index;
    }


    string PathCovert(string path)
    {
        //如果是绝对路径不进行转化
        if (path.Contains(":") || File.Exists(path))
        {
            return path;
        }
        else
        {
            //查找路径
            //先从文档路径开始找
            //找到后再查找对应Data目录

            List<string> files = FileTool.GetAllFileNamesByPath(PathDefine.GetDataPath(), new string[] { "txt" });

            for (int i = 0; i < files.Count; i++)
            {
                string fileName = FileTool.RemoveExpandName(FileTool.GetFileNameByPath(files[i]));

                if (fileName == path)
                {
                    return files[i];
                }
            }

            MessageBox.Show("没有找到 " + path);

            return null;
        }
    }

    public bool GetFileIsExist()
    {
        //如果是绝对路径不进行转化
        if (m_txtName.Contains(":")  )
        {
            return File.Exists(m_txtName);
        }
        else
        {
            List<string> files = FileTool.GetAllFileNamesByPath(PathDefine.GetDataPath(), new string[] { "txt" });

            for (int i = 0; i < files.Count; i++)
            {
                string fileName = FileTool.RemoveExpandName(FileTool.GetFileNameByPath(files[i]));

                if (fileName == m_txtName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
