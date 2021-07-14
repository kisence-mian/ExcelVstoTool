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

    static Dictionary<string, DataTable> dataCache = new Dictionary<string, DataTable>();

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
}
