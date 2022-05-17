using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

    public static bool isFirstInit = true;

    public static void Init()
    {
        string configPath = PathDefine.GetLanguageDataPath() + @"/" + Const.c_ConfigName_Language;

        if(File.Exists(configPath))
        {
            IsEnable = true;

            string configJson = FileTool.ReadStringByFile(configPath);
            LanguageConfig config = JsonConvert.DeserializeObject<LanguageConfig>(configJson);

            currentLanguage = config.gameExistLanguages[0];
            allLanuage = config.gameExistLanguages;

            //清空缓存
            GenerateLanguageCache();
        }
        else
        {
            IsEnable = false;

            //只在第一次初始化时弹出这个提示
            if (isFirstInit)
            {
                MessageBox.Show("没有找到 " + configPath);
            }
        }

        if(isFirstInit)
        {
            isFirstInit = false;
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

            if(!languageCache.ContainsKey(language))
            {
                languageCache.Add(language, dict);
            }
            else
            {
                MessageBox.Show("检测到重复的多语言设置" + language + " 请检查 LanguageConfig.txt");
            }
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

    /// <summary>
    /// 判断一个字符串是不是多语言Key
    /// </summary>
    /// <returns></returns>
    public static bool GetIsLanguageKey(string value)
    {
        //目前的判断条件是：
        //含有字符 ‘/’
        //且文件名与字段名都不为空

        if(value.Contains('/'))
        {
            if( !string.IsNullOrEmpty(GetFileName(value))
                &&!string.IsNullOrEmpty(GetLanguageKey(value)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        else
        {
            return false;
        }
    }

    public static string GetFileName(string value)
    {
        string fileName = "";

        //选择数组中的第一个打开
        if(value.Contains("|"))
        {
            value = value.Split('|')[0];
        }

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

            case SystemLanguage.Portuguese:
                return "pt";
        }


        return ((int)language).ToString();
    }

    public static List<string> GetLanguageFileNameList(SystemLanguage language)
    {
        List<string> list = new List<string>();

        //允许空值
        list.Add("");

        if(!languageCache.ContainsKey(language))
        {
            return list;
        }

        foreach (string key in languageCache[language].Keys)
        {
            //排除其他空值
            if(!string.IsNullOrEmpty(key))
            {
                list.Add(key);
            }
        }

        return list;
    }

    public static List<string> GetLanguageKeyList(SystemLanguage language, string fileName)
    {
        List<string> list = new List<string>();

        for (int i = 0; i < languageCache[language][fileName].TableIDs.Count; i++)
        {
            string id = languageCache[language][fileName].TableIDs[i];
            string languageKey = (fileName + "_" + id).Replace("_","/");

            list.Add(languageKey);
        }

        return list;
    }

    public static string GetLanguageContent(SystemLanguage language, string languageKey)
    {
        string pattern = "{#[0-9a-zA-Z/_]+}";

        string fileName = GetFileName(languageKey);
        string key = GetLanguageKey(languageKey); ;
        try
        {
            SingleData sData = languageCache[language][fileName][key];
            string value = sData.GetString("value");

            //查找是否有嵌套Key
            if(Regex.IsMatch(value, pattern))
            {
                var matachs = Regex.Matches(value, pattern);

                for (int i = 0; i < matachs.Count; i++)
                {
                    string subKey = matachs[i].Value.Replace("{#", "").Replace("}","");
                    value = value.Replace(matachs[i].Value, GetLanguageContent(language, subKey));
                }
            }

            return value;
        }
        catch (Exception e)
        {
            return "" + e.Message;
        }
    }

    public static void CreateLanguageFile(SystemLanguage dataLanguage,string fileName,Dictionary<string,string> languageData)
    {
        for (int i = 0; i < allLanuage.Count; i++)
        {
            SystemLanguage language = allLanuage[i];

            CreateSingleLanguageFile(language, fileName, languageData, language == dataLanguage);
        }
    }

    static void CreateSingleLanguageFile(SystemLanguage language, string fileName, Dictionary<string, string> languageData,bool isFull)
    {
        DataTable data = new DataTable();
        data.TableKeys.Add(Const.c_LanguageData_mainKey);
        data.TableKeys.Add(Const.c_LanguageData_valueKey);
        data.SetDefault(Const.c_LanguageData_valueKey, Const.c_LanguageData_defaultValue);

        foreach (var item in languageData)
        {
            SingleData sd = new SingleData();
            sd.Add(Const.c_LanguageData_mainKey, item.Key);

            if(isFull)
            {
                sd.Add(Const.c_LanguageData_valueKey, item.Value);
            }
            data.AddData(sd);
        }

        //写入缓存
        languageCache[language].Add(fileName,data);

        string filePath = PathDefine.GetLanguageDataPath() +@"/" + language.ToString() + @"/" + Const.c_LanguagePrefix + "_" + language.ToString() + "_" + fileName + ".txt";
        FileTool.WriteStringByFile(filePath,DataTable.Serialize(data));
    }
}
