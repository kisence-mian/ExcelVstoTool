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
    #region 常量

    public const string c_ConfigSheetName = "Config";
    public const string c_AssetsDireName = "Assets";

    #endregion

    public string m_sheetName;
    public string m_txtPath;

    //导入时覆盖Excel公式
    public bool m_coverFormula;

    public DataConfig(Worksheet configSheet, int index)
    {
        string key             = "A" + index;
        string key_txtPath     = "B" + index;
        string key_saveFormula = "C" + index;

        m_sheetName = configSheet.Range[key].Value;
        m_txtPath = PathCovert(configSheet.Range[key_txtPath].Value);

        m_coverFormula = (bool)ParseTool.GetBool( configSheet.Range[key_saveFormula].Value);
    }

    public static void ConfigInit(Worksheet configSheet)
    {
        configSheet.Range["A1"].Value = "页签名称";
        configSheet.Range["B1"].Value = "文件名";
        configSheet.Range["C1"].Value = "覆盖公式";
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

            string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
            string assetsPath = filePath.Substring(0, filePath.IndexOf(c_AssetsDireName) + c_AssetsDireName.Length);
            string dataPath = assetsPath + @"\Resources\Data";

            List<string> files = FileTool.GetAllFileNamesByPath(dataPath, new string[] { "txt" });

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
}
