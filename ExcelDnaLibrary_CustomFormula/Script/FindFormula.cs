using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula.Script
{
    public class FindFormula
    {
        #region 查找
        [ExcelFunction(Name = "GetValueByDefault", Description = "如果目标位置为空值，则返回默认值", Category = "查找")]
        public static object GetValueByDefault(
             [ExcelArgument(Name = "Value", Description = "目标值")]Object value,
             [ExcelArgument(Name = "DefaultValue", Description = "默认值")]Object defaultValue)
        {
            if (value.GetType() != typeof(ExcelEmpty))
            {
                return defaultValue;
            }
            else
            {
                return value;
            }
        }

        [ExcelFunction(Name = "FindData", Description = "以指定位置的值去进行匹配，并输出结果", Category = "查找")]
        public static object FindData(
            [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
            [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
            [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]Object[] valueRange
            )
        {
            try
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();

                for (int i = 0; i < keyRange.Length; i++)
                {
                    if (valueRange.Length > i)
                    {
                        string keyTemp = ObjectToString(keyRange[i]);
                        if (!string.IsNullOrEmpty(keyTemp) && !dict.ContainsKey(keyTemp))
                        {
                            dict.Add(keyTemp, valueRange[i]);
                        }
                    }
                }

                if (dict.ContainsKey(key))
                {
                    return dict[key];
                }
                else
                {
                    return "Dont Find Value by:->" + key + "<-";
                }
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "IsFindData", Description = "返回目标范围是否存在指定的Key", Category = "查找")]
        public static bool IsFindData(
            [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
            [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange
            )
        {
            Dictionary<string, bool> dict = new Dictionary<string, bool>();

            for (int i = 0; i < keyRange.Length; i++)
            {
                string keyTemp = ObjectToString(keyRange[i]);
                if (!string.IsNullOrEmpty(keyTemp) && !dict.ContainsKey(keyTemp))
                {
                    dict.Add(keyTemp, true);
                }
            }

            if (dict.ContainsKey(key))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        [ExcelFunction(Name = "FindData_Average", Description = "以指定位置的值去进行匹配，并输出所有匹配结果的平均数", Category = "查找")]
        public static object FindData_Average(
        [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
        [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
        [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]double[] valueRange
        )
        {
            try
            {
                Dictionary<string, List<double>> dict = new Dictionary<string, List<double>>();

                //构造数据
                for (int i = 0; i < keyRange.Length; i++)
                {
                    if (valueRange.Length > i)
                    {
                        string keyTemp = ObjectToString(keyRange[i]);
                        if (!string.IsNullOrEmpty(keyTemp))
                        {
                            List<double> list;

                            if (dict.ContainsKey(keyTemp))
                            {
                                list = dict[keyTemp];
                            }
                            else
                            {
                                list = new List<double>();
                                dict.Add(keyTemp, list);
                            }

                            list.Add(valueRange[i]);
                        }
                    }
                }

                if (dict.ContainsKey(key))
                {
                    //取平均数
                    List<double> list = dict[key];

                    double sum = 0;

                    for (int i = 0; i < list.Count; i++)
                    {
                        sum += list[i];
                    }

                    return sum / list.Count;
                }
                else
                {
                    return "Dont Find Value by:->" + key + "<-";
                }
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "FindData_Sum", Description = "以指定位置的值去进行匹配，并输出所有匹配结果的和", Category = "查找")]
        public static object FindData_Sum(
        [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
        [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
        [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]double[] valueRange
)
        {
            try
            {
                Dictionary<string, List<double>> dict = new Dictionary<string, List<double>>();

                //构造数据
                for (int i = 0; i < keyRange.Length; i++)
                {
                    if (valueRange.Length > i)
                    {
                        string keyTemp = ObjectToString(keyRange[i]);
                        if (!string.IsNullOrEmpty(keyTemp))
                        {
                            List<double> list;

                            if (dict.ContainsKey(keyTemp))
                            {
                                list = dict[keyTemp];
                            }
                            else
                            {
                                list = new List<double>();
                                dict.Add(keyTemp, list);
                            }

                            list.Add(valueRange[i]);
                        }
                    }
                }

                if (dict.ContainsKey(key))
                {
                    //取和
                    List<double> list = dict[key];

                    double sum = 0;

                    for (int i = 0; i < list.Count; i++)
                    {
                        sum += list[i];
                    }

                    return sum;
                }
                else
                {
                    return "Dont Find Value by:->" + key + "<-";
                }
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "FindData_Median", Description = "以指定位置的值去进行匹配，并输出所有匹配结果的中位数", Category = "查找")]
        public static object FindData_Median(
        [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
        [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
        [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]double[] valueRange
)
        {
            try
            {
                Dictionary<string, List<double>> dict = new Dictionary<string, List<double>>();

                //构造数据
                for (int i = 0; i < keyRange.Length; i++)
                {
                    if (valueRange.Length > i)
                    {
                        string keyTemp = ObjectToString(keyRange[i]);
                        if (!string.IsNullOrEmpty(keyTemp))
                        {
                            List<double> list;

                            if (dict.ContainsKey(keyTemp))
                            {
                                list = dict[keyTemp];
                            }
                            else
                            {
                                list = new List<double>();
                                dict.Add(keyTemp, list);
                            }

                            list.Add(valueRange[i]);
                        }
                    }
                }

                if (dict.ContainsKey(key))
                {
                    //取中位数
                    List<double> list = dict[key];
                    list.Sort();

                    return list[list.Count / 2];
                }
                else
                {
                    return "Dont Find Value by:->" + key + "<-";
                }
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "FindData_Max", Description = "以指定位置的值去进行匹配，并输出所有匹配结果的最大值", Category = "查找")]
        public static object FindData_Max(
           [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
           [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
           [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]double[] valueRange
)
        {
            try
            {
                Dictionary<string, List<double>> dict = new Dictionary<string, List<double>>();

                //构造数据
                for (int i = 0; i < keyRange.Length; i++)
                {
                    if (valueRange.Length > i)
                    {
                        string keyTemp = ObjectToString(keyRange[i]);
                        if (!string.IsNullOrEmpty(keyTemp))
                        {
                            List<double> list;

                            if (dict.ContainsKey(keyTemp))
                            {
                                list = dict[keyTemp];
                            }
                            else
                            {
                                list = new List<double>();
                                dict.Add(keyTemp, list);
                            }

                            list.Add(valueRange[i]);
                        }
                    }
                }

                if (dict.ContainsKey(key))
                {
                    //取最大值
                    List<double> list = dict[key];
                    double max = double.MinValue;

                    for (int i = 0; i < list.Count; i++)
                    {
                        if (list[i] > max)
                        {
                            max = list[i];
                        }
                    }

                    return max;
                }
                else
                {
                    return "Dont Find Value by:->" + key + "<-";
                }
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "FindData_Min", Description = "以指定位置的值去进行匹配，并输出所有匹配结果的最小值", Category = "查找")]
        public static object FindData_Min(
          [ExcelArgument(Name = "Key", Description = "要查询的Key")]string key,
          [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
          [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]double[] valueRange
)
        {
            try
            {
                Dictionary<string, List<double>> dict = new Dictionary<string, List<double>>();

                //构造数据
                for (int i = 0; i < keyRange.Length; i++)
                {
                    if (valueRange.Length > i)
                    {
                        string keyTemp = ObjectToString(keyRange[i]);
                        if (!string.IsNullOrEmpty(keyTemp))
                        {
                            List<double> list;

                            if (dict.ContainsKey(keyTemp))
                            {
                                list = dict[keyTemp];
                            }
                            else
                            {
                                list = new List<double>();
                                dict.Add(keyTemp, list);
                            }

                            list.Add(valueRange[i]);
                        }
                    }
                }

                if (dict.ContainsKey(key))
                {
                    //取最小值
                    List<double> list = dict[key];
                    double min = double.MaxValue;

                    for (int i = 0; i < list.Count; i++)
                    {
                        if (list[i] < min)
                        {
                            min = list[i];
                        }
                    }

                    return min;
                }
                else
                {
                    return "Dont Find Value by:->" + key + "<-";
                }
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayFindData", Description = "传入一个数组，按照指定位置去匹配结果，返回以分隔符连接的数组", Category = "查找")]
        public static string ArrayFindData(
            [ExcelArgument(Name = "Array", Description = "要查询的Key Array")]string keyArray,
            [ExcelArgument(Name = "KeyRange", Description = "匹配Key的范围")]Object[] keyRange,
            [ExcelArgument(Name = "ValueRange", Description = "输出结果的范围")]Object[] valueRange,
             [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            string sc = splitChar;
            if (sc == "|")
            {
                sc = @"\|";
            }

            string[] f = Regex.Split(keyArray, sc, RegexOptions.IgnoreCase);

            string result = "";

            for (int i = 0; i < f.Length; i++)
            {
                result += FindData(f[i], keyRange, valueRange);

                if (i != f.Length - 1)
                {
                    result += splitChar;
                }
            }

            return result;
        }

        [ExcelFunction(Name = "FindNotNullData", Description = "返回范围内第一个不为空的数据", Category = "查找")]
        public static object FindOrdinalNotNullData(
            [ExcelArgument(Name = "Range", Description = "要查询的范围")]Object[] range)
        {
            for (int i = 0; i < range.Length; i++)
            {
                if (range[i].GetType() != typeof(ExcelEmpty))
                {
                    return range[i];
                }
            }

            return "Dont Find not empty value";
        }

        [ExcelFunction(Name = "FindByRegex_Is", Description = "返回范围内第一个与正则表达式匹配的项", Category = "查找")]
        public static bool IsFindByRegex(
        [ExcelArgument(Name = "Range", Description = "要查询的范围")]Object[] range,
        [ExcelArgument(Name = "Pattern", Description = "要查询的范围")]string pattern)
        {
            for (int i = 0; i < range.Length; i++)
            {
                string content = range[i].ToString();

                if (range[i].GetType() == typeof(ExcelEmpty))
                {
                    content = "";
                }

                if (Regex.IsMatch(content, pattern))
                {
                    return true;
                }
            }

            return false;
        }

        [ExcelFunction(Name = "FindByRegex", Description = "返回范围内第一个与正则表达式匹配的项", Category = "查找")]
        public static object FindByRegex(
        [ExcelArgument(Name = "Range", Description = "要查询的范围")]Object[] range,
        [ExcelArgument(Name = "Pattern", Description = "要查询的范围")]string pattern)
        {
            for (int i = 0; i < range.Length; i++)
            {
                string content = range[i].ToString();

                if (range[i].GetType() == typeof(ExcelEmpty))
                {
                    content = "";
                }

                if (Regex.IsMatch(content, pattern))
                {
                    return range[i];
                }
            }

            return "Dont Find by Regex" + pattern;
        }

        [ExcelFunction(Name = "FindOrdinalNotNullData", Description = "返回范围内第N个不为空的数据", Category = "查找")]
        public static object FindOrdinalNotNullData(
            [ExcelArgument(Name = "Range", Description = "要查询的范围")]Object[] range,
            [ExcelArgument(Name = "Index", Description = "数目")]int index = 1)
        {
            int temp = 0;
            for (int i = 0; i < range.Length; i++)
            {
                if (range[i].GetType() != typeof(ExcelEmpty))
                {
                    temp++;
                }

                if (temp == index)
                {
                    return range[i];
                }
            }

            return "Dont Find by " + index;
        }

        [ExcelFunction(Name = "FindOrdinalNotNullData_Index", Description = "返回范围内第N个不为空的数据,注意空值会返回0", Category = "查找")]
        public static int FindOrdinalNotNullData_Index(
            [ExcelArgument(Name = "Range", Description = "要查询的范围")]Object[] range,
            [ExcelArgument(Name = "Index", Description = "数目")]int index)
        {
            int temp = 0;
            for (int i = 0; i < range.Length; i++)
            {
                if (range[i].GetType() != typeof(ExcelEmpty))
                {
                    temp++;
                }

                if (temp == index)
                {
                    return i + 1;
                }
            }

            return -1;
        }

        static string ObjectToString(object obj)
        {
            if (obj == null)
            {
                return "";
            }
            else if (obj is ExcelEmpty)
            {
                return "";
            }
            else
            {
                return obj.ToString();
            }
        }

        #endregion
    }
}
