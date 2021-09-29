using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDnaLibrary_CustomFormula
{

    public class CustomFormula
    {
        [ExcelFunction(Name = "TextSplit", Description ="将字符串按指定的格式分割，并取出指定位置的项",Category ="字符串处理")]
        public static string TextSplit(
            
            [ExcelArgument(Name = "Content",Description = "要分割的文本")]string content,
            [ExcelArgument(Name = "SplitChar",Description ="分隔符")]string splitChar,
            [ExcelArgument(Name = "Index",Description ="序号")]int index)
        {
            try
            {
                if(splitChar == "|")
                {
                    splitChar = @"\|";
                }

                string[] array = Regex.Split(content, splitChar, RegexOptions.IgnoreCase);

                if(index >=0 && index < array.Length)
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


        //[ExcelFunction(Name = "test")]
        //public static string Test([ExcelArgument(Name = "Content")]string content)
        //{
        //    return "im test->" + content;
        //}
    }
}
