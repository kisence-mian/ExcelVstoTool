using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

public static class DataManager
{
    /// <summary>
    /// 是否生效
    /// </summary>
    public static bool IsEnable = false;

    /// <summary>
    /// 是否选中在了有效区域内
    /// </summary>
    public static bool isWorkRange = false;

    private static FieldType currentFieldType;
    private static DataFieldAssetType currentAssetType;
    private static string currentSecType ="";

    /// <summary>
    /// 外部数据缓存
    /// </summary>
    static Dictionary<string, DataTable> dataCache = new Dictionary<string, DataTable>();

    /// <summary>
    /// 表名缓存
    /// </summary>
    static List<string> tableNameCache = new List<string>();

    static List<string> enumNameCache = new List<string>();

    public static FieldType CurrentFieldType
    {
        get => currentFieldType;
        set
        {
            //重设其他属性的值
            if (value != currentFieldType)
            {
                currentAssetType = DataFieldAssetType.Data;

                if(value == FieldType.Enum && enumNameCache.Count > 0)
                {
                    currentSecType = enumNameCache[0];
                }
                else
                {
                    currentSecType = "";
                }
            }

            currentFieldType = value;
        }
    }

    public static DataFieldAssetType CurrentAssetType
    {
        get => currentAssetType;
        set
        {
            //重设二级类型的值
            if (currentAssetType != value)
            {
                if(value == DataFieldAssetType.TableKey && tableNameCache.Count>0)
                {
                    currentSecType = tableNameCache[0];
                }
                else
                {
                    currentSecType = "";
                }
            }

            currentAssetType = value;
        }
    }

    public static string CurrentSecType { get => currentSecType; set => currentSecType = value; }
    public static List<string> TableName { get => tableNameCache;  }
    public static List<string> EnumName { get => enumNameCache;  }

    public static void Init()
    {
        IsEnable = true;

        //清空缓存
        GenerateLanguageCache();
    }

    static void GenerateLanguageCache()
    {
        List<string> errorFile = new List<string>();
        dataCache = new Dictionary<string, DataTable>();

        List<string> list = FileTool.GetAllFileNamesByPath(PathDefine.GetDataPath(), new string[] { "txt" },false);

        for (int j = 0; j < list.Count; j++)
        {
            string key = FileTool.RemoveExpandName(FileTool.GetFileNameByPath(list[j]));
            try
            {
                string content = FileTool.ReadStringByFile(list[j]);
                DataTable data = DataTable.Analysis(content);

                dataCache.Add(key, data);
            }
            catch (Exception)
            {
                errorFile.Add(key);
            }
        }

        //缓存表名
        tableNameCache = new List<string>();
        foreach (string key in dataCache.Keys)
        {
            tableNameCache.Add(key);
        }

        //TODO 构造枚举名

        if (errorFile.Count > 0)
        {
            string info = "构造以下配置表时出现错误，请检查相应文件";
            for (int i = 0; i < errorFile.Count; i++)
            {
                info += "\n->" + errorFile[i];
            }

            MessageBox.Show(info);
        }
    }

    public static bool CheckDataFileNameExist( string fileName)
    {
        if (!IsEnable)
            return true;

        return dataCache.ContainsKey(fileName);
    }

    public static bool CheckDataExist( string fileName, string key)
    {
        if (!IsEnable)
            return true;

        return dataCache[fileName].ContainsKey(key);
    }

    public static List<string> GetTableKeyList(string tableName)
    {
        List<string> list = new List<string>();

        if(dataCache.ContainsKey(tableName))
        {
            return dataCache[tableName].TableIDs;
        }

        return list;
    }


    /// <summary>
    /// 解析
    /// </summary>
    /// <param name="typeString"></param>
    public static void PaseToCurrentType(string typeString)
    {
        if(typeString == null)
        {
            isWorkRange = false;
        }
        else
        {
            isWorkRange = true;

            if(typeString == "")
            {
                //全部默认值
                CurrentFieldType = FieldType.String;
                CurrentAssetType = DataFieldAssetType.Data;
                CurrentSecType = "";
            }
            else
            {
                string[] tempType = typeString.Split(DataTable.c_DataFieldAssetTypeSplit);
                string[] content = tempType[0].Split(DataTable.c_EnumSplit);

                try
                {
                    string fieldType = content[0];
                    if (fieldType.Contains("["))
                    {
                        string[] tempSS = fieldType.Split('[');
                        fieldType = tempSS[0];
                        string splitStr = tempSS[1].Replace("]", "");
                        //data.m_ArraySplitFormat.Add(field, splitStr.ToCharArray());
                    }

                    CurrentFieldType = (FieldType)Enum.Parse(typeof(FieldType), fieldType);

                    if (content.Length > 1)
                    {
                        CurrentSecType = content[1];
                    }
                    else
                    {
                        CurrentSecType = "";
                    }
                }
                catch (Exception)
                {
                    CurrentFieldType = FieldType.String;
                    CurrentSecType = "";
                }

                if (tempType.Length > 1)
                {
                    CurrentAssetType =(DataFieldAssetType)Enum.Parse(typeof(DataFieldAssetType), tempType[1]);
                }
                else
                {
                    CurrentAssetType = DataFieldAssetType.Data;
                }
            }
        }
    }

    public static string GetCurrentTypeString()
    {
        string typeString = CurrentFieldType.ToString();

        if(!string.IsNullOrEmpty(CurrentSecType))
        {
            typeString += "|" + CurrentSecType;
        }

        if(CurrentAssetType != DataFieldAssetType.Data)
        {
            typeString += "&" + CurrentAssetType;
        }

        return typeString;
    }
}
