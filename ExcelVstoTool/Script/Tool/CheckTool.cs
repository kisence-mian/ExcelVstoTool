using ExcelVstoTool;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class CheckTool
{
    public static DataTable CheckSheet(Worksheet workSheet,DataConfig config)
    {
        DataTable data = null;

        try
        {
            //表头校验
            CheckTitle(workSheet, config);

            data = DataTool.Excel2Table(workSheet, config);

            //ID校验
            CheckID(workSheet, data, config);

            //类型校验
            CheckType(workSheet, data, config);

            //格式校验
            CheckFormat(workSheet, data, config);

            //外部校验
            CheckResource(workSheet, data, config);

        }
        catch (Exception e)
        {
            System.Windows.Forms.MessageBox.Show("校验出错 ->" + config.m_sheetName + " \n" + e.Message /*+ "\n" + e.ToString()*/);
            return null;
        }
        return data;
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

    static bool CheckID(Worksheet workSheet, DataTable data, DataConfig config)
    {
        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];

            //检查ID中的空格
            if (id.Contains(" "))
            {
                throw new Exception("ID 含有空格 >" + id + "<\n  第 " + (i + 5) + " 行 ");
            }
        }

        return true;
    }

    static bool CheckType(Worksheet workSheet, DataTable data, DataConfig config)
    {
        bool result = true;

        //读取类型
        foreach (string key in data.TableKeys)
        {
            //跳过主键的类型
            if (key == data.TableKeys[0])
            {
                continue;
            }

            if (data.m_tableTypes.ContainsKey(key))
            {
                FieldType type = data.m_tableTypes[key];

                switch (type)
                {
                    case FieldType.Enum:
                    case FieldType.EnumArray:
                        string enumType = data.m_tableSecTypes[key];
                        result &= CheckEnum(workSheet, data, config, key, enumType);
                        break;
                }
            }

        }
        return result;
    }

    static bool CheckFormat(Worksheet workSheet, DataTable data, DataConfig config)
    {
        bool result = true;

        //读取类型
        foreach (string key in data.TableKeys)
        {
            //跳过主键的类型
            if (key == data.TableKeys[0])
            {
                continue;
            }

            for (int i = 0; i < data.TableIDs.Count; i++)
            {
                string id = data.TableIDs[i];
                SingleData sData = data[id];

                if (data.m_tableTypes.ContainsKey(key))
                {
                    FieldType assetType = data.m_tableTypes[key];
                    string value = sData.GetString(key);

                    try
                    {
                        switch (assetType)
                        {
                            case FieldType.Bool: sData.GetBool(key); break;
                            case FieldType.BoolArray: sData.GetBoolArray(key); break;

                            case FieldType.Int: sData.GetInt(key); break;
                            case FieldType.IntArray: sData.GetIntArray(key); break;

                            case FieldType.Float: sData.GetFloat(key); break;
                            case FieldType.FloatArray: sData.GetFloatArray(key); break;

                            case FieldType.Vector2: sData.GetVector2(key); break;
                            case FieldType.Vector2Array: sData.GetVector2Array(key); break;

                            case FieldType.Vector3: sData.GetVector3(key); break;
                            case FieldType.Vector3Array: sData.GetVector3Array(key); break;

                            case FieldType.Color: sData.GetColor(key); break;

                            default: break;
                        }
                    }
                    catch (Exception e)
                    {
                        throw new Exception("格式不匹配 ID = " + id + " 第 " + (i + 5) + "行 Key:" + key
                            + "\n 类型是" + assetType + " --->" + value + "<-"
                            + CheckSpace(sData.GetString(key))
                            + "\n" + e.ToString());
                    }
                }
                else
                {
                    throw new Exception("找不到  ->" + key + "<- 的类型 行 2");
                }
            }

        }
        return result;
    }

    //校验外部资源
    static bool CheckResource(Worksheet workSheet, DataTable data, DataConfig config)
    {
        bool result = true;

        //读取类型
        foreach(string key in data.TableKeys)
        {
            if(data.m_fieldAssetTypes.ContainsKey(key))
            {
                DataFieldAssetType assetType = data.m_fieldAssetTypes[key];
                FieldType type = data.m_tableTypes[key];

                switch (assetType)
                {
                    case DataFieldAssetType.Prefab:
                        result &= CheckPerfab(workSheet,data,config,key, type);
                        break;
                    case DataFieldAssetType.Texture:
                        result &= CheckTexture(workSheet, data, config, key, type);
                        break;
                    case DataFieldAssetType.Audio:
                        result &= CheckAudio(workSheet, data, config, key, type);
                        break;
                    case DataFieldAssetType.Video:
                        result &= CheckVideo(workSheet, data, config, key, type);
                        break;
                    case DataFieldAssetType.LocalizedLanguage:
                        result &= CheckLanguage(workSheet, data, config, key, type);
                        break;
                    case DataFieldAssetType.TableName:
                        result &= CheckTableName(workSheet, data, config, key, type);
                        break;
                    case DataFieldAssetType.TableKey:

                        if(data.m_tableSecTypes.ContainsKey(key))
                        {
                            string tableKey = data.m_tableSecTypes[key];
                            result &= CheckTableKey(workSheet, data, config, key, tableKey, type);
                        }
                        else
                        {
                            throw new Exception("字段" + key + " 没有指定配置表名称 ");
                        }

                        break;


                    default:break;
                }
            }
        }

        return result;
    }

    static bool CheckPerfab(Worksheet workSheet, DataTable data, DataConfig config,string key,FieldType fieldType)
    {
        bool result = true;

        //构造预设清单
        List<string> list = FileTool.GetAllFileNamesByPath(PathDefine.GetResourcesPath(), new string[] { "prefab" });
        Dictionary<string, string> dict = GenerateNameDict(list);

        //逐项检查表格中的数据是否存在

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[id];
            string value = sData.GetString(key);

            if (!CheckSingleResource(fieldType, sData, key, (v) => {
                return dict.ContainsKey(v);
            }))
            {
                throw new Exception("找不到 预设资源 -> " + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                    + CheckSpace(sData.GetString(key)));
            }
        }

        return result;
    }

    static bool CheckTexture(Worksheet workSheet, DataTable data, DataConfig config, string key, FieldType fieldType)
    {
        bool result = true;

        //构造图片清单
        List<string> list = FileTool.GetAllFileNamesByPath(PathDefine.GetResourcesPath(), new string[] { "png","jpg","jpeg" });
        Dictionary<string, string> dict = GenerateNameDict(list);

        //逐项检查表格中的数据是否存在

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[id];
            string value = sData.GetString(key);

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }

            if (!CheckSingleResource(fieldType, sData, key, (v)=>{
                return dict.ContainsKey(v);
            }))
            {
                throw new Exception("找不到 图片资源 -> " + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                    + CheckSpace(sData.GetString(key)));
            }
        }
        return result;
    }

    static bool CheckAudio(Worksheet workSheet, DataTable data, DataConfig config, string key, FieldType fieldType)
    {
        bool result = true;

        //构造图片清单
        List<string> list = FileTool.GetAllFileNamesByPath(PathDefine.GetResourcesPath(), new string[] { "mp3", "wav", "mid" , "ogg" });
        Dictionary<string, string> dict = GenerateNameDict(list);

        //逐项检查表格中的数据是否存在

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[id];
            string value = sData.GetString(key);

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }

            if (!CheckSingleResource(fieldType, sData, key, (v) => {
                return dict.ContainsKey(v);
            }))
            {
                throw new Exception("找不到 音频资源 -> " + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                    + CheckSpace(sData.GetString(key)));
            }
        }
        return result;
    }

    static bool CheckVideo(Worksheet workSheet, DataTable data, DataConfig config, string key, FieldType fieldType)
    {
        bool result = true;

        //构造视频清单
        List<string> list = FileTool.GetAllFileNamesByPath(PathDefine.GetResourcesPath(), new string[] { "mp4", "mkv", "avi" , "flv" });
        Dictionary<string, string> dict = GenerateNameDict(list);

        //逐项检查表格中的数据是否存在

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[id];
            string value = sData.GetString(key);

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }

            if (!CheckSingleResource(fieldType, sData, key, (v) => {
                return dict.ContainsKey(v);
            }))
            {
                throw new Exception("找不到 音频资源 -> " + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                    + CheckSpace(sData.GetString(key)));
            }
        }
        return result;
    }

    static bool CheckSingleResource(FieldType fieldType, SingleData sData, string key, CheckHandle handle)
    {
        if (fieldType == FieldType.String
            || fieldType == FieldType.Enum)
        {
            string value = sData.GetString(key);

            //忽略空值
            if(string.IsNullOrEmpty(value))
            {
                return true;
            }

            return handle(value);
        }
        else if (fieldType == FieldType.StringArray
            || fieldType == FieldType.EnumArray)
        {
            string[] values = sData.GetStringArray(key);

            for (int i = 0; i < values.Length; i++)
            {
                //忽略空值
                if (string.IsNullOrEmpty(values[i]))
                {
                    return true;
                }

                if (!handle(values[i]))
                {
                    return false;
                }
            }

            return true;
        }
        else
        {
            throw new Exception("意料之外的类型 " + fieldType);
        }
    }

    static bool CheckLanguage(Worksheet workSheet, DataTable data, DataConfig config, string key, FieldType fieldType)
    {
        bool result = true;

        if(!LanguageManager.IsEnable)
        {
            return true;
        }

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[data.TableIDs[i]];
            string value = sData.GetString(key);

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }

            for (int j = 0; j < LanguageManager.allLanuage.Count; j++)
            {
                SystemLanguage language = LanguageManager.allLanuage[j];

                string fn = LanguageManager.GetFileName(value);
                string lk = LanguageManager.GetLanguageKey(value);

                if (!CheckSingleResource(fieldType, sData, key, (v) =>
                {
                    string fileName = LanguageManager.GetFileName(v);
                    string languageKey = LanguageManager.GetLanguageKey(v);
                    return LanguageManager.CheckLanguageFileNameExist(language, fileName);
                }))
                {
                    throw new Exception("多语言文件错误 ->" + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                        + "\n找不到 多语言文件 "+ fn
                        + "\n语种是 " + language
                        + CheckSpace(value));
                }

                //if (!LanguageManager.CheckLanguageExist(language, fileName, languageKey))
                if (!CheckSingleResource(fieldType, sData, key, (v) =>
                {
                    string fileName = LanguageManager.GetFileName(v);
                    string languageKey = LanguageManager.GetLanguageKey(v);

                    return LanguageManager.CheckLanguageExist(language, fileName, languageKey);
                }))
                {
                    throw new Exception("多语言Key错误 ->" + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                       + "\n找不到 多语言Key " + lk
                       + "\n语种是 " + language
                       + CheckSpace(value));
                }
            }
        }

        return result;
    }

    static bool CheckTableKey(Worksheet workSheet, DataTable data, DataConfig config, string key, string tableKey, FieldType fieldType)
    {
        bool result = true;

        if (!DataManager.IsEnable)
        {
            return true;
        }

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[data.TableIDs[i]];
            string value = sData.GetString(key);

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }


            if (!CheckSingleResource(fieldType, sData, key, (v) =>
            {
                return DataManager.CheckDataFileNameExist(tableKey);
            }))
            {
                throw new Exception("配置表Key错误 ->" + tableKey + "<- 行 2 ID=" + id + " Key " + key
                    + "\n找不到 配置表文件 " + tableKey
                    + CheckSpace(tableKey));
            }

            if (!CheckSingleResource(fieldType, sData, key, (v) =>
            {
                bool r = DataManager.CheckDataExist(tableKey, v);

                if(!r)
                {
                    throw new Exception("配置表Key错误 ->" + value + "<- 行 " + (i + 5) + " ID = " + id + " Key =" + key
                   + "\n找不到 配置表Key " + v
                   + CheckSpace(v));
                }

                return DataManager.CheckDataExist(tableKey, v);
            }))
            {
                throw new Exception("配置表Key错误 ->" + value + "<- 行 " + (i + 5) + " ID = " + id + " Key " + key
                   + "\n找不到 配置表Key " + value
                   + CheckSpace(value));
            }
        }

        return result;
    }

    static bool CheckTableName(Worksheet workSheet, DataTable data, DataConfig config, string key, FieldType fieldType)
    {
        bool result = true;

        if (!DataManager.IsEnable)
        {
            return true;
        }

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[data.TableIDs[i]];
            string value = sData.GetString(key);

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }

            if (!CheckSingleResource(fieldType, sData, key, (v) =>
            {
                return DataManager.CheckDataFileNameExist(v);
            }))
            {
                throw new Exception("配置表Key错误 ->" + value + "<- 行 " + (i + 5) + " ID=" + id + " Key " + key
                    + "\n找不到 配置表文件 " + value
                    + CheckSpace(value));
            }
        }

        return result;
    }

    static bool CheckEnum(Worksheet workSheet, DataTable data, DataConfig config, string key, string enumName)
    {
        bool result = true;

        if (!DataManager.IsEnable)
        {
            return true;
        }

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[data.TableIDs[i]];
            string value = sData.GetString(key);
            FieldType fieldType = data.m_tableTypes[key];

            //跳过空数据
            if (string.IsNullOrEmpty(value))
            {
                continue;
            }

            if (!CheckSingleResource(fieldType, sData, key, (v) =>
            {
                return DataManager.GetEnumList(enumName).Contains(v);
            }))
            {
                throw new Exception("找不到枚举 ->" + value + "<- 行 " + (i + 5) + " ID=" + id + " Key " + key
                    + "\n枚举名称 " + enumName
                    + CheckSpace(value));
            }
        }

        return result;
    }

    static string CheckSpace(string content)
    {
        if(content == null)
        {
            return "";
        }

        if (content.Contains(" "))
        {
            return "\n注意文本里有空格";
        }
        else
        {
            return "";
        }
    }

    static Dictionary<string,string> GenerateNameDict(List<string> list)
    {
        Dictionary<string, string> dict = new Dictionary<string, string>();

        for (int i = 0; i < list.Count; i++)
        {
            string fileName = FileTool.RemoveExpandName( FileTool.GetFileNameByPath(list[i]));

            dict.Add(fileName, list[i]);
        }
        return dict;
    }

    public delegate bool CheckHandle(string value);
}
