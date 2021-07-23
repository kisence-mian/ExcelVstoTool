using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

public class DataTool
{
    public static void Excel2Data(Worksheet excel, DataConfig dataConfig)
    {
        DataTable data = new DataTable();

        Worksheet _wsh = excel;

        try
        {
            data = Excel2Table(excel, dataConfig);
            FileTool.WriteStringByFile(dataConfig.GetTextPath(), DataTable.Serialize(data));
        }
        catch (Exception e)
        {
            System.Windows.Forms.MessageBox.Show(dataConfig.GetTextPath() + "->" + e.ToString());
        }
    }

    public static DataTable Excel2Table(Worksheet excel, DataConfig dataConfig)
    {
        DataTable data = new DataTable();
        Worksheet _wsh = excel;

        //解析key
        int col = 1;
        int row = 1;

        try
        {
            int totalCol = 0;


            while (!string.IsNullOrEmpty(_wsh.Cells[row, col].Text.ToString()))
            {
                data.TableKeys.Add(_wsh.Cells[row, col].Text.ToString());
                col++;
            }

            totalCol = col - 1;

            //解析类型
            row = 2;
            string[] lineData = new string[totalCol];
            for (col = 1; col <= totalCol; col++)
            {
                lineData[col - 1] = _wsh.Cells[row, col].Text.ToString();
            }

            DataTable.AnalysisFieldType(data, lineData);

            //解析注释
            row = 3;
            lineData = new string[totalCol];
            for (col = 1; col <= totalCol; col++)
            {
                lineData[col - 1] = _wsh.Cells[row, col].Text.ToString();
            }

            DataTable.AnalysisNoteValue(data, lineData);

            //解析默认值
            row = 4;
            lineData = new string[totalCol];
            for (col = 1; col <= totalCol; col++)
            {
                lineData[col - 1] = _wsh.Cells[row, col].Text.ToString();
            }

            DataTable.AnalysisDefaultValue(data, lineData);

            //解析值
            row = 5;
            col = 1;
            while (!string.IsNullOrEmpty(_wsh.Cells[row, 1].Text.ToString()))
            {
                //Console.WriteLine("wsh.Cells["+col+", " + row + "] = " + _wsh.Cells[col, 1].Text.ToString());

                SingleData dataTmp = new SingleData();
                dataTmp.data = data;
                for (col = 1; col <= totalCol; col++)
                {
                    if(!string.IsNullOrEmpty(_wsh.Cells[row, col].Text.ToString()))
                    {
                        dataTmp.Add(data.TableKeys[col - 1], _wsh.Cells[row, col].Text.ToString());
                    }
                }
                row++;

                data.AddData(dataTmp);
            }
        }
        catch (Exception e)
        {
            throw new Exception("解析->" + dataConfig.m_sheetName +" 错误! 行 " + row + "\n错误内容 " + e.Message);
        }

        return data;
    }

    public static void Data2Excel(DataConfig dataConfig, Worksheet excel)
    {
        DataTable data = DataTable.Analysis(FileTool.ReadStringByFile(dataConfig.GetTextPath()));

        Worksheet _wsh = excel;

        //如果全覆盖公式则整个删除
        ExcelTool.ClearSheet(_wsh,dataConfig.m_coverFormula);

        try
        {
            //表头
            int index = 1;
            foreach (var key in data.TableKeys)
            {
                _wsh.Cells[1, index] = key;

                index++;
            }

            //字段类型
            index = 1;
            foreach (var key in data.TableKeys)
            {
                string typeString = "";
                if (data.m_tableTypes.ContainsKey(key))
                {
                    typeString = data.m_tableTypes[key].ToString();

                    if (data.m_ArraySplitFormat.ContainsKey(key))
                    {
                        typeString += "[";
                        foreach (var item in data.m_ArraySplitFormat[key])
                        {
                            typeString += item;
                        }
                        typeString += "]";
                    }

                    if (data.m_tableSecTypes.ContainsKey(key))
                    {
                        typeString += "|" + data.m_tableSecTypes[key];
                    }

                    if (data.m_fieldAssetTypes.ContainsKey(key))
                    {
                        if (data.m_fieldAssetTypes[key] != DataFieldAssetType.Data)
                            typeString += "&" + data.m_fieldAssetTypes[key];
                    }

                    _wsh.Cells[2, index] = typeString;
                }
                else
                {
                    _wsh.Cells[2, index] = FieldType.String.ToString();
                }

                index++;
            }
            _wsh.Cells[2, 1] = DataTable.c_fieldTypeTableTitle;

            //描述
            index = 1;
            _wsh.Cells[3, 1] = DataTable.c_noteTableTitle;
            foreach (var key in data.TableKeys)
            {
                if(data.m_noteValue.ContainsKey(key))
                {
                    _wsh.Cells[3, index] = data.m_noteValue[key];
                }

                index++;
            }

            //默认值
            index = 1;
            _wsh.Cells[4, 1] = DataTable.c_defaultValueTableTitle;
            foreach (var key in data.TableKeys)
            {
                if (data.m_defaultValue.ContainsKey(key))
                {
                    _wsh.Cells[4, index] = data.m_defaultValue[key];
                }
                index++;
            }

            //值
            int row = 5;
            foreach (var id in data.TableIDs)
            {
                index = 1;
                SingleData dataTmp = data[id];
                foreach (var key in data.TableKeys)
                {
                    if (dataTmp.ContainsKey(key))
                    {
                        //读取是否覆盖公式的设置
                        if(_wsh.Cells[row, index].HasFormula)
                        {
                            if(dataConfig.m_coverFormula)
                            {
                                _wsh.Cells[row, index] = dataTmp[key].ToString();
                            }
                            else
                            {
                                //跳过公式覆盖
                            }
                        }
                        else
                        {
                            _wsh.Cells[row, index] = dataTmp[key].ToString();
                        }
                    }
                    index++;
                }
                row++;
            }


        }
        catch (Exception e)
        {
            MessageBox.Show("导入出错  "+ dataConfig .m_txtName+ "\n" + e.ToString());
        }
    }

    public static void CreateNewData(string filePath)
    {
        DataTable data = new DataTable();
        data.TableKeys.Add(Const.c_LanguageData_mainKey);
        data.TableKeys.Add(Const.c_LanguageData_valueKey);
        data.SetDefault(Const.c_LanguageData_valueKey, "这里填写默认值");
        data.SetNote(Const.c_LanguageData_valueKey, "这里填写描述");

        FileTool.WriteStringByFile(filePath, DataTable.Serialize(data));
    }

    #region 自动生成代码

    /// <summary>
    /// 创建数据表对应的具体数据类
    /// </summary>
    /// <param name="dataName"></param>
    /// <param name="data"></param>
    public static string CreateDataCSharpFile(string dataName, DataTable data)
    {
        if (dataName.Contains("/"))
        {
            string[] tmp = dataName.Split('/');
            dataName = tmp[tmp.Length - 1];
        }

        string className = dataName + "Generate";
        string content = "";

        content += "using System;\n";
        content += "using UnityEngine;\n\n";

        content += @"//" + className + "类\n";
        content += @"//该类自动生成请勿修改，以避免不必要的损失";
        content += "\n";

        content += "public class " + className + " : DataGenerateBase \n";
        content += "{\n";

        content += "\tpublic string m_key;\n";

        //type
        List<string> type = new List<string>(data.m_tableTypes.Keys);

        //Debug.Log("type count: " + type.Count);

        if (type.Count > 0)
        {
            for (int i = 1; i < data.TableKeys.Count; i++)
            {
                string key = data.TableKeys[i];
                string enumType = null;

                if (data.m_tableSecTypes.ContainsKey(key))
                {
                    enumType = data.m_tableSecTypes[key];
                }
                char[] m_ArraySplitFormat = new char[0];
                if (data.m_ArraySplitFormat.ContainsKey(key))
                {
                    m_ArraySplitFormat = data.m_ArraySplitFormat[key];
                }

                if (data.m_noteValue.ContainsKey(key))
                {
                    content += "\t" + @"/// <summary>" + "\n";
                    content += "\t" + @"/// " + data.m_noteValue[key] + "\n";
                    content += "\t" + @"/// </summary>" + "\n";
                }

                content += "\t";

                if (data.m_tableTypes.ContainsKey(key))
                {
                    //访问类型 + 字段类型  + 字段名
                    content += "public " + OutPutFieldName(data.m_tableTypes[key], enumType, m_ArraySplitFormat) + " m_" + key + ";";
                }
                //默认字符类型
                else
                {
                    //访问类型 + 字符串类型 + 字段名 
                    content += "public " + "string" + " m_" + key + ";";
                }

                content += "\n\n";
            }
        }

        content += "\n";

        content += "\tpublic override void LoadData(string key) \n";
        content += "\t{\n";
        content += "\t\tDataTable table =  DataManager.GetData(\"" + dataName + "\");\n\n";
        content += "\t\tif (!table.ContainsKey(key))\n";
        content += "\t\t{\n";
        content += "\t\t\tthrow new Exception(\"" + className + " LoadData Exception Not Fond key ->\" + key + \"<-\");\n";
        content += "\t\t}\n";
        content += "\n";
        content += "\t\tSingleData data = table[key];\n\n";

        content += "\t\tm_key = key;\n";

        if (type.Count > 0)
        {
            for (int i = 1; i < data.TableKeys.Count; i++)
            {
                string key = data.TableKeys[i];

                content += "\t\t";


                string enumType = null;

                if (data.m_tableSecTypes.ContainsKey(key))
                {
                    enumType = data.m_tableSecTypes[key];
                }
                char[] m_ArraySplitFormat = new char[0];
                if (data.m_ArraySplitFormat.ContainsKey(key))
                {
                    m_ArraySplitFormat = data.m_ArraySplitFormat[key];
                }

                if (data.m_tableTypes.ContainsKey(key))
                {
                    content += "m_" + key + " = data." + OutPutFieldFunction(data.m_tableTypes[key], enumType, m_ArraySplitFormat) + "(\"" + key + "\")";
                }
                //默认字符类型
                else
                {
                    content += "m_" + key + " = data." + OutPutFieldFunction(FieldType.String, enumType, m_ArraySplitFormat) + "(\"" + key + "\")";
                    //Debug.LogWarning("字段 " + key + "没有配置类型！");
                }

                content += ";\n";
            }
        }

        content += "\t}\n";
        content += "\t public override void LoadData(DataTable table,string key) \n";
        content += "\t{\n";

        content += "\t\tSingleData data = table[key];\n\n";

        content += "\t\tm_key = key;\n";

        if (type.Count > 0)
        {
            for (int i = 1; i < data.TableKeys.Count; i++)
            {
                string key = data.TableKeys[i];

                content += "\t\t";

                string enumType = null;

                if (data.m_tableSecTypes.ContainsKey(key))
                {
                    enumType = data.m_tableSecTypes[key];
                }
                char[] m_ArraySplitFormat = new char[0];
                if (data.m_ArraySplitFormat.ContainsKey(key))
                {
                    m_ArraySplitFormat = data.m_ArraySplitFormat[key];
                }
                if (data.m_tableTypes.ContainsKey(key))
                {
                    content += "m_" + key + " = data." + OutPutFieldFunction(data.m_tableTypes[key], enumType, m_ArraySplitFormat) + "(\"" + key + "\")";
                }
                //默认字符类型
                else
                {
                    content += "m_" + key + " = data." + OutPutFieldFunction(FieldType.String, enumType, m_ArraySplitFormat) + "(\"" + key + "\")";
                    //Debug.LogWarning("字段 " + key + "没有配置类型！");
                }

                content += ";\n";
            }
        }

        content += "\t}\n";

        content += "}\n";

        return content;

        //string SavePath = Application.dataPath + "/Script/DataClassGenerate/" + className + ".cs";

        //EditorUtil.WriteStringByFile(SavePath, content.ToString());
    }



    static string OutPutFieldFunction(FieldType fileType, string enumType, char[] m_ArraySplitFormat)
    {
        string arrayFun = "";
        for (int i = 0; i < m_ArraySplitFormat.Length; i++)
        {
            arrayFun += "[]";
        }


        switch (fileType)
        {
            case FieldType.Bool: return "GetBool";
            case FieldType.Color: return "GetColor";
            case FieldType.Float: return "GetFloat";
            case FieldType.Int: return "GetInt";
            case FieldType.String: return "GetString";
            case FieldType.Vector2: return "GetVector2";
            case FieldType.Vector3: return "GetVector3";
            case FieldType.Enum: return "GetEnum<" + enumType + ">";

            case FieldType.StringArray:
                arrayFun = "string" + arrayFun;
                break;
            case FieldType.IntArray:
                arrayFun = "int" + arrayFun;
                break;
            case FieldType.FloatArray:
                arrayFun = "float" + arrayFun;
                break;
            case FieldType.BoolArray:
                arrayFun = "bool" + arrayFun;
                break;
            case FieldType.Vector2Array:
                arrayFun = "Vector2" + arrayFun;
                break;
            case FieldType.Vector3Array:
                arrayFun = "Vector3" + arrayFun;
                break;

        }
        arrayFun = "GetArray<" + arrayFun + ">";
        return arrayFun;
    }

    static string OutPutFieldName(FieldType fileType, string enumType, char[] m_ArraySplitFormat)
    {
        string arrayFun = "";
        for (int i = 0; i < m_ArraySplitFormat.Length; i++)
        {
            arrayFun += "[]";
        }
        switch (fileType)
        {
            case FieldType.Bool: return "bool";
            case FieldType.Color: return "Color";
            case FieldType.Float: return "float";
            case FieldType.Int: return "int";
            case FieldType.String: return "string";
            case FieldType.Vector2: return "Vector2";
            case FieldType.Vector3: return "Vector3";
            case FieldType.Enum: return enumType;

            case FieldType.StringArray: return "string[]" + arrayFun;
            case FieldType.IntArray: return "int[]" + arrayFun;
            case FieldType.FloatArray: return "float[]" + arrayFun;
            case FieldType.BoolArray: return "bool[]" + arrayFun;
            case FieldType.Vector2Array: return "Vector2[]" + arrayFun;
            case FieldType.Vector3Array: return "Vector3[]" + arrayFun;
            default: return "";
        }
    }

    #endregion
}

