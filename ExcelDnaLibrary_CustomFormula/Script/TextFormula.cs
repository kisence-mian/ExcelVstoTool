using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula.Script
{
    public class TextFormula
    {
        #region 字符串处理

        [ExcelFunction(Name = "TextSplit", Description = "将字符串按指定的格式分割，并取出指定位置的项,支持使用正则式作为分隔符", Category = "字符串处理")]
        public static string TextSplit(

            [ExcelArgument(Name = "Content", Description = "要分割的文本")]string content,
            [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar,
            [ExcelArgument(Name = "Index", Description = "序号")]int index)
        {
            try
            {
                if (splitChar == "|")
                {
                    splitChar = @"\|";
                }

                string[] array = Regex.Split(content, splitChar, RegexOptions.IgnoreCase);

                if (index >= 0 && index < array.Length)
                {
                    return array[index];
                }
                else
                {
                    return "Split Error: array Length is " + array.Length + " index is " + index;
                }
            }
            catch (Exception)
            {
                return "Error";
            }
        }

        [ExcelFunction(Name = "TextSplitByDefaultLast", Description = "将字符串按指定的格式分割，并取出指定位置的项,支持使用正则式作为分隔符,如果分隔后的结果不够，则默认取最后一个", Category = "字符串处理")]
        public static string TextSplitByDefaultLast(

            [ExcelArgument(Name = "Content", Description = "要分割的文本")]string content,
            [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar,
            [ExcelArgument(Name = "Index", Description = "序号")]int index)
        {
            try
            {
                if (splitChar == "|")
                {
                    splitChar = @"\|";
                }

                string[] array = Regex.Split(content, splitChar, RegexOptions.IgnoreCase);

                if (index >= 0 && index < array.Length)
                {
                    return array[index];
                }
                else
                {
                    index = array.Length - 1;
                    return array[index];
                }
            }
            catch (Exception)
            {
                return "Error";
            }
        }


        [ExcelFunction(Name = "TextSplitLength", Description = "将字符串按指定的格式分割，返回分隔的份数", Category = "字符串处理")]
        public static int GetArrayLength(
            [ExcelArgument(Name = "Content", Description = "要分割的数组文本")]string content,
            [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] array = Regex.Split(content, splitChar, RegexOptions.IgnoreCase);

            return array.Length;
        }

        [ExcelFunction(Name = "BatchText", Description = "将字符串复制n份，用分隔符连接，并将{i}替换为序号", Category = "字符串处理")]
        public static string BatchText(
            [ExcelArgument(Name = "Content", Description = "要复制的文本")]string content,
            [ExcelArgument(Name = "Count", Description = "复制次数")]int count,
            [ExcelArgument(Name = "StartIndex", Description = "初始序号")]int startIndex,
            [ExcelArgument(Name = "LinkChar", Description = "连接符")]string linkChar)
        {
            string result = "";

            for (int i = 0; i < count; i++)
            {
                result += content.Replace("{i}", (i + startIndex) + "");

                if (i != count - 1)
                {
                    result += linkChar;
                }
            }

            return result;
        }

        [ExcelFunction(Name = "TextReplace", Description = "进行文本替换，支持使用正则表达式", Category = "字符串处理")]
        public static string RegexReplace(
        [ExcelArgument(Name = "Input", Description = "原始文本")]string input,
        [ExcelArgument(Name = "Pattern", Description = "正则公式")]string pattern,
        [ExcelArgument(Name = "Replacement", Description = "替换内容")]string replacement)
        {
            return Regex.Replace(input, pattern, replacement);
        }

        [ExcelFunction(Name = "TextRegexIsMatch", Description = "使用正则表达式进行匹配，返回是或者否", Category = "字符串处理")]
        public static bool RegexIsMatch(
        [ExcelArgument(Name = "Input", Description = "原始文本")]string input,
        [ExcelArgument(Name = "Pattern", Description = "正则公式")]string pattern)
        {
            return Regex.IsMatch(input, pattern);
        }

        [ExcelFunction(Name = "TextRegexMatch", Description = "使用正则表达式进行匹配，返回匹配结果", Category = "字符串处理")]
        public static string RegexMatch(
        [ExcelArgument(Name = "Input", Description = "原始文本")]string input,
        [ExcelArgument(Name = "Pattern", Description = "正则公式")]string pattern)
        {
            return Regex.Match(input, pattern).Value;
        }

        #endregion

    }
}
