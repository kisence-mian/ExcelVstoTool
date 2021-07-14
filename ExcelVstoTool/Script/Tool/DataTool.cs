using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


public class DataTool
{
    public static void Excel2Data(Worksheet excel, DataConfig dataConfig)
    {
        DataTable data = new DataTable();

        Worksheet _wsh = excel;

        try
        {
            data = Excel2Table(excel, dataConfig);
            FileTool.WriteStringByFile(dataConfig.m_txtPath, DataTable.Serialize(data));
        }
        catch (Exception e)
        {
            System.Windows.Forms.MessageBox.Show(dataConfig.m_txtPath + "->" + e.ToString());
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
        DataTable data = DataTable.Analysis(FileTool.ReadStringByFile(dataConfig.m_txtPath));

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
                _wsh.Cells[4, index] = data.m_defaultValue[key];
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
            Console.WriteLine(dataConfig.m_txtPath + "->" + e.ToString());
        }
    }
}

