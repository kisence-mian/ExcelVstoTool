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
        #region 数列

        [ExcelFunction(Name = "AddSelfArray", Description = "构造一个用连接符连接的自增(等差)数列", Category = "数列")]
        public static string AddSelfArray(
          [ExcelArgument(Name = "Start", Description = "起始数字")]int start,
          [ExcelArgument(Name = "Change", Description = "自增数字")]int change,
          [ExcelArgument(Name = "Count", Description = "数列长度")]int count,
          [ExcelArgument(Name = "LinkChar", Description = "连接符")]string linkChar)
        {
            string result = "";
            int number = start;

            for (int i = 0; i < count; i++)
            {
                result += number;

                number += change;

                if (i != count - 1)
                {
                    result += linkChar;
                }
            }


            return result;
        }

        [ExcelFunction(Name = "MulSelfArray", Description = "构造一个用连接符连接的等比数列", Category = "数列")]
        public static string MulSelfArray(
            [ExcelArgument(Name = "Start", Description = "起始数字")]int start,
            [ExcelArgument(Name = "Mul", Description = "变化比值")]int mul,
            [ExcelArgument(Name = "Count", Description = "数列长度")]int count,
            [ExcelArgument(Name = "LinkChar", Description = "连接符")]string linkChar)
        {
            string result = "";
            int number = start;

            for (int i = 0; i < count; i++)
            {
                result += number;

                number *= mul;

                if (i != count - 1)
                {
                    result += linkChar;
                }
            }

            return result;
        }

        #endregion

        #region 计算

        [ExcelFunction(Name = "ArrayAdd", Description = "将两个数组的数字相加，并将结果相加，不匹配的长度会被舍弃", Category = "数组")]
        public static string ArrayAdd(
           [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
           [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
           [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] fA = Regex.Split(arrayA, splitChar, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, splitChar, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            float result = 0;
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) + float.Parse(fB[i]);
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayAdd2Array", Description = "将两个数组的数字相加，返回一个新的字符串数组", Category = "数组")]
        public static string ArrayAdd2Array(
            [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
            [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
            [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            string sc = splitChar;
            if (sc == "|")
            {
                sc = @"\|";
            }

            string[] fA = Regex.Split(arrayA, sc, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, sc, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            string result = "";
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) + float.Parse(fB[i]);

                    if (i != fA.Length - 1)
                    {
                        result += splitChar;
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }


        [ExcelFunction(Name = "ArraySub", Description = "将两个数组的数字相减(A-B)，并将结果相加，不匹配的长度会被舍弃", Category = "数组")]
        public static string ArraySub(
            [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
            [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
            [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] fA = Regex.Split(arrayA, splitChar, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, splitChar, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            float result = 0;
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) + float.Parse(fB[i]);
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArraySubArray", Description = "将两个数组的数字相减（A - B），返回一个新的字符串数组", Category = "数组")]
        public static string ArraySub2Array(
            [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
            [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
            [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            string sc = splitChar;
            if (sc == "|")
            {
                sc = @"\|";
            }

            string[] fA = Regex.Split(arrayA, sc, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, sc, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            string result = "";
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) - float.Parse(fB[i]);

                    if (i != fA.Length - 1)
                    {
                        result += splitChar;
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }


        [ExcelFunction(Name = "ArrayMul", Description = "将两个数组的数字相乘，并将结果相加，不匹配的长度会被舍弃", Category = "数组")]
        public static string ArrayMul(
         [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
         [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] fA = Regex.Split(arrayA, splitChar, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, splitChar, RegexOptions.IgnoreCase);

            if(fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            float result = 0;
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) * float.Parse(fB[i]);
                }

                return result.ToString();
            }
            catch(Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayMul2Array", Description = "将两个数组的数字相乘，并将结果相加，不匹配的长度会被舍弃,返回一个新的字符串数组", Category = "数组")]
        public static string ArrayMul2Array(
         [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
         [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            string sc = splitChar;
            if (sc == "|")
            {
                sc = @"\|";
            }

            string[] fA = Regex.Split(arrayA, sc, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, sc, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            string result = "";
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) * float.Parse(fB[i]);

                    if(i != fA.Length - 1)
                    {
                        result += splitChar;
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayDiv", Description = "将两个数组的数字相除 (A除以B)，并将结果相加，不匹配的长度会被舍弃", Category = "数组")]
        public static string ArrayDiv(
         [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
         [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] fA = Regex.Split(arrayA, splitChar, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, splitChar, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            float result = 0;
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) / float.Parse(fB[i]);
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayDiv2Array", Description = "将两个数组的数字相除 (A除以B)，并将结果相加，不匹配的长度会被舍弃,返回一个新的字符串数组", Category = "数组")]
        public static string ArrayDiv2Array(
         [ExcelArgument(Name = "Array A", Description = "数组 A")]string arrayA,
         [ExcelArgument(Name = "Array B", Description = "数组 B")]string arrayB,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            string sc = splitChar;
            if (sc == "|")
            {
                sc = @"\|";
            }

            string[] fA = Regex.Split(arrayA, sc, RegexOptions.IgnoreCase);
            string[] fB = Regex.Split(arrayB, sc, RegexOptions.IgnoreCase);

            if (fA.Length != fB.Length)
            {
                return "no mactch length ! A.Length = " + fA.Length + " B.Length = " + fB.Length;
            }

            string result = "";
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    result += float.Parse(fA[i]) / float.Parse(fB[i]);

                    if (i != fA.Length - 1)
                    {
                        result += splitChar;
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayMax", Description = "返回数组中的最大值", Category = "数组")]
        public static string ArrayMax(
         [ExcelArgument(Name = "Array", Description = "数组")]string array,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] f = Regex.Split(array, splitChar, RegexOptions.IgnoreCase);

            float result = float.MinValue;
            try
            {
                for (int i = 0; i < f.Length; i++)
                {
                    if(float.Parse(f[i]) > result)
                    {
                        result = float.Parse(f[i]);
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayMin", Description = "返回数组中的最大值", Category = "数组")]
        public static string ArrayMin(
         [ExcelArgument(Name = "Array", Description = "数组")]string array,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] f = Regex.Split(array, splitChar, RegexOptions.IgnoreCase);

            float result = float.MaxValue;
            try
            {
                for (int i = 0; i < f.Length; i++)
                {
                    if (float.Parse(f[i]) < result)
                    {
                        result = float.Parse(f[i]);
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArraySum", Description = "返回数组中的和", Category = "数组")]
        public static string ArraySum(
         [ExcelArgument(Name = "Array", Description = "数组")]string array,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] f = Regex.Split(array, splitChar, RegexOptions.IgnoreCase);

            float result = 0;
            try
            {
                for (int i = 0; i < f.Length; i++)
                {
                    result += float.Parse(f[i]);
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArraySumByIndex", Description = "返回数组前几项的和", Category = "数组")]
        public static string ArraySumByIndex(
         [ExcelArgument(Name = "Array", Description = "数组")]string array,
         [ExcelArgument(Name = "Index", Description = "第几项")]int index,
         [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            if (splitChar == "|")
            {
                splitChar = @"\|";
            }

            string[] f = Regex.Split(array, splitChar, RegexOptions.IgnoreCase);

            float result = 0;

            if(index > f.Length)
            {
                index = f.Length;
            }

            try
            {
                for (int i = 0; i < index; i++)
                {
                    result += float.Parse(f[i]);
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        [ExcelFunction(Name = "ArrayNormal2Array", Description = "将目标数组归一（正数变成1，负数变为-1）,返回一个新的字符串数组", Category = "数组")]
        public static string ArrayNormal2Array(
           [ExcelArgument(Name = "Array A", Description = "数组 A")]string array,
           [ExcelArgument(Name = "SplitChar", Description = "分隔符")]string splitChar)
        {
            string sc = splitChar;
            if (sc == "|")
            {
                sc = @"\|";
            }

            string[] fA = Regex.Split(array, sc, RegexOptions.IgnoreCase);

            string result = "";
            try
            {
                for (int i = 0; i < fA.Length; i++)
                {
                    float temp = float.Parse(fA[i]);

                    if(temp > 0)
                    {
                        temp = 1;
                    }
                    else if( temp  < 0)
                    {
                        temp = -1;
                    }
                    else
                    {
                        temp = 0;
                    }

                    result += temp;

                    if (i != fA.Length - 1)
                    {
                        result += splitChar;
                    }
                }

                return result.ToString();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
        }

        #endregion
    }
}
