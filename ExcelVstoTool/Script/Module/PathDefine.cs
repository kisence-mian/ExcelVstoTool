using ExcelVstoTool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public static class PathDefine
{
    public static bool IsAssetsPath()
    {
        string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
        return filePath.Contains(Const.c_DireName_Assets);
    }


    public static string GetResourcesPath()
    {
        string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
        string assetsPath = filePath.Substring(0, filePath.IndexOf(Const.c_DireName_Assets) + Const.c_DireName_Assets.Length);
        string resourcesPath = assetsPath + @"\" + Const.c_DireName_Resources;

        return resourcesPath;
    }

    public static string GetDataPath()
    {
        string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
        string assetsPath = filePath.Substring(0, filePath.IndexOf(Const.c_DireName_Assets) + Const.c_DireName_Assets.Length);
        string resourcesPath = assetsPath + @"\" + Const.c_DireName_Resources + @"\" + Const.c_DireName_Data;

        return resourcesPath;
    }

    public static string GetLanguageDataPath()
    {
        string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
        string assetsPath = filePath.Substring(0, filePath.IndexOf(Const.c_DireName_Assets) + Const.c_DireName_Assets.Length);
        string resourcesPath = assetsPath + @"\" + Const.c_DireName_Resources + @"\" + Const.c_DireName_Data + @"\" + Const.c_DireName_Language;

        return resourcesPath;
    }

    public static string GetLanguageDataPath(SystemLanguage language)
    {
        string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
        string assetsPath = filePath.Substring(0, filePath.IndexOf(Const.c_DireName_Assets) + Const.c_DireName_Assets.Length);
        string resourcesPath = assetsPath + @"\" + Const.c_DireName_Resources + @"\" + Const.c_DireName_Data + @"\" + Const.c_DireName_Language + @"\" + language.ToString();

        return resourcesPath;
    }

    public static string GetDataGeneratePath()
    {
        string filePath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
        string assetsPath = filePath.Substring(0, filePath.IndexOf(Const.c_DireName_Assets) + Const.c_DireName_Assets.Length);
        string resourcesPath = assetsPath + @"\" + Const.c_DireName_Script + @"\" + Const.c_DireName_DataClassGenerate;

        return resourcesPath;
    }
}
