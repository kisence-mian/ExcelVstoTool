using ExcelVstoTool;
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

            //类型校验
            DataTable data = DataTool.Excel2Table(workSheet, config);

            //格式校验

            //外部校验
            result &= CheckResource(workSheet, data, config);

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

                switch(assetType)
                {
                    case DataFieldAssetType.Prefab:
                        result &= CheckPerfab(workSheet,data,config,key);
                        break;
                    case DataFieldAssetType.Texture:
                        result &= CheckTexture(workSheet, data, config, key);
                        break;
                    case DataFieldAssetType.LocalizedLanguage:
                        result &= CheckLanguage(workSheet, data, config, key);
                        break;
                    case DataFieldAssetType.TableKey:

                        if(data.m_tableSecTypes.ContainsKey(key))
                        {
                            string tableKey = data.m_tableSecTypes[key];
                            result &= CheckTableKey(workSheet, data, config, key, tableKey);
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

    static bool CheckPerfab(Worksheet workSheet, DataTable data, DataConfig config,string key)
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
            if (!dict.ContainsKey(sData.GetString(key)))
            {
                throw new Exception("找不到 预设资源 -> " + sData.GetString(key) + "<- 行 " + (i + 5) + " ID = " + id
                    + CheckSpace(sData.GetString(key))); 
            }
        }

        return result;
    }

    static bool CheckTexture(Worksheet workSheet, DataTable data, DataConfig config, string key)
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

            if (!dict.ContainsKey(sData.GetString(key)))
            {
                throw new Exception("找不到 图片资源 -> " + sData.GetString(key)  + "<- 行 " + (i + 5) + " ID = " + id 
                    + CheckSpace(sData.GetString(key)));
            }
        }
        return result;
    }

    static bool CheckLanguage(Worksheet workSheet, DataTable data, DataConfig config, string key)
    {
        bool result = true;

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[data.TableIDs[i]];
            string value = sData.GetString(key);

            string fileName = LanguageManager.GetFileName(value);
            string languageKey = LanguageManager.GetLanguageKey(value);

            for (int j = 0; j < LanguageManager.allLanuage.Count; j++)
            {
                SystemLanguage language = LanguageManager.allLanuage[j];

                if(!LanguageManager.CheckLanguageFileNameExist(language, fileName))
                {
                    throw new Exception("多语言Key错误 ->" + value + "<- 行 " + (i + 5) + " ID = " + id
                        + "\n找不到 多语言文件 "+ fileName
                        + CheckSpace(value));
                }

                if (!LanguageManager.CheckLanguageExist(language, fileName, languageKey))
                {
                    throw new Exception("多语言Key错误 ->" + value + "<- 行 " +(i + 5) + " ID = " + id
                       + "\n找不到 多语言Key " + languageKey
                       + CheckSpace(value));
                }
            }
        }

        return result;
    }

    static bool CheckTableKey(Worksheet workSheet, DataTable data, DataConfig config, string key,string tableKey)
    {
        bool result = true;

        for (int i = 0; i < data.TableIDs.Count; i++)
        {
            string id = data.TableIDs[i];
            SingleData sData = data[data.TableIDs[i]];
            string value = sData.GetString(key);

            if (!DataManager.CheckDataFileNameExist(tableKey))
            {
                throw new Exception("配置表Key错误 ->" + tableKey + "<- 行 2 ID=" + id
                    + "\n找不到 配置表文件 " + tableKey
                    + CheckSpace(tableKey));
            }

            if (!DataManager.CheckDataExist(tableKey, value))
            {
                throw new Exception("配置表Key错误 ->" + value + "<- 行 " + (i + 5) + " ID = " + id
                   + "\n找不到 配置表Key " + value
                   + CheckSpace(value));
            }
        }

        return result;
    }

    static string CheckSpace(string content)
    {
        if(content.Contains(" "))
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


}
