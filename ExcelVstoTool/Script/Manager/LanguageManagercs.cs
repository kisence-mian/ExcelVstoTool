using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

public static class LanguageManager
{
    /// <summary>
    /// 是否生效
    /// </summary>
    public static bool IsEnable = false;

    /// <summary>
    /// 当前使用的多语言
    /// </summary>
    public static SystemLanguage currentLanguage;

    /// <summary>
    /// 所有可用的多语言
    /// </summary>
    public static List<SystemLanguage> allLanuage;

    /// <summary>
    /// 多语言缓存
    /// </summary>
    static Dictionary<SystemLanguage, Dictionary<string, DataTable>> languageCache = new Dictionary<SystemLanguage, Dictionary<string, DataTable>>();

    public static void Init()
    {
        string configPath = PathDefine.GetLanguageDataPath() + @"/" + Const.c_ConfigName_Language;

        if(File.Exists(configPath))
        {
            IsEnable = true;

            string configJson = FileTool.ReadStringByFile(configPath);
            LanguageConfig config = JsonConvert.DeserializeObject<LanguageConfig>(configJson);

            currentLanguage = config.defaultLanguage;
            allLanuage = config.gameExistLanguages;

            //清空缓存
            GenerateLanguageCache();
        }
        else
        {
            IsEnable = false;

            MessageBox.Show("没有找到 " + configPath);
        }
    }

    static void GenerateLanguageCache()
    {
        List<string> errorFile = new List<string>();
        languageCache = new Dictionary<SystemLanguage, Dictionary<string, DataTable>>();

        for (int i = 0; i < allLanuage.Count; i++)
        {
            SystemLanguage language = allLanuage[i];
            Dictionary<string, DataTable> dict = new Dictionary<string, DataTable>();

            List<string> list = FileTool.GetAllFileNamesByPath(PathDefine.GetLanguageDataPath(language), new string[] { "txt" });

            for (int j = 0; j < list.Count; j++)
            {
                string key = FileTool.RemoveExpandName(FileTool.GetFileNameByPath(list[j]));
                try
                {
                    string simpleKey = key.Replace(Const.c_LanguagePrefix +"_" + language.ToString() + "_", "");

                    string content = FileTool.ReadStringByFile(list[j]);
                    DataTable data = DataTable.Analysis(content);

                    dict.Add(simpleKey, data);
                }
                catch (Exception) 
                {
                    errorFile.Add(key);
                }
            }

            languageCache.Add(language, dict);
        }

        if(errorFile.Count > 0)
        {
            string info = "构造以下多语言表时出现错误，请检查相应文件";
            for (int i = 0; i < errorFile.Count; i++)
            {
                info += "\n->" + errorFile[i];
            }

            MessageBox.Show(info);
        }
    }

    public static string GetFileName(string value)
    {
        string fileName = "";
        string[] path = value.Split('/');

        for (int i = 0; i < path.Length -1; i++)
        {
            fileName += path[i];

            if(i != path.Length - 2)
            {
                fileName += "_";
            }
        }

        return fileName;
    }

    public static string GetLanguageKey(string value)
    {
        string[] path = value.Split('/');

        return path[path.Length - 1];
    }

    public static bool CheckLanguageFileNameExist(SystemLanguage language,string fileName)
    {
        if(languageCache.ContainsKey(language))
        {
            return languageCache[language].ContainsKey(fileName);
        }
        else
        {
            throw new Exception("找不到对应的语言类型 " + language);
        }
    }

    public static bool CheckLanguageExist(SystemLanguage language, string fileName, string key)
    {
        if (languageCache.ContainsKey(language))
        {
            return languageCache[language][fileName].ContainsKey(key);
        }
        else
        {
            throw new Exception("找不到对应的语言类型 " + language);
        }
    }

    public static string GetLanguageAcronym(SystemLanguage language)
    {
        switch(language)
        {
            case SystemLanguage.Chinese:
            case SystemLanguage.ChineseSimplified:
                return "cn";

            case SystemLanguage.ChineseTraditional:
                return "hk";

            case SystemLanguage.English:
                return "en";

            case SystemLanguage.Japanese:
                return "jp";

            case SystemLanguage.Russian:
                return "ru";
        }


        return ((int)language).ToString();
    }

    public static string GetLanguageContent(SystemLanguage language, string languageKey)
    {
        string fileName = GetFileName(languageKey);
        string key = GetLanguageKey(languageKey); ;
        try
        {
            SingleData sData = languageCache[language][fileName][key];
            return sData.GetString("value");
        }
        catch (Exception e)
        {
            return "" + e.Message;
        }
    }
}
