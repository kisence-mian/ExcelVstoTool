using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula.Script
{
    public class TimeFormula
    {
        static string s_dtFormat = "yyyy-MM-dd HH:mm:ss";

        [ExcelFunction(Name = "Time_UnixTime2Date",
            Description = "将Unix时间戳转化为指定日期 输出格式为 yyyy-MM-dd hh:mm:ss", Category = "时间")]

        public static string Time_UnixTime2Date(
            [ExcelArgument(Name = "UnixTime", Description = "要分割的文本")]string input)
        {
            // 用来格式化long类型时间的,声明的变量
            long unixDate;
            DateTime start;
            DateTime date;
            //ENd

            unixDate = long.Parse(input);
            start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            date = start.AddMilliseconds(unixDate).ToLocalTime();

            DateTime dtStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            DateTime newTime = dtStart.AddSeconds(unixDate);

            return newTime.ToString(s_dtFormat);
        }

        [ExcelFunction(Name = "Time_GetTimeOffet",
            Description = "计算两个时间的差值 输出格式为 yyyy-MM-dd hh:mm:ss", Category = "时间")]
        public static string Time_GetTimeOffet(
           [ExcelArgument(Name = "UnixTime", Description = "要分割的文本")]string dateA,
            [ExcelArgument(Name = "UnixTime", Description = "要分割的文本")]string dateB)
        {
            try
            {
                DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
                dtFormat.ShortDatePattern = s_dtFormat;

                DateTime a = Convert.ToDateTime(dateA, dtFormat);
                DateTime b = Convert.ToDateTime(dateB, dtFormat);

                var offset = a - b;

                return offset.ToString();

            }
            catch(Exception e)
            {
                return "Error: A:" + dateA + " B:" + dateB + " \n  Message:" + e.Message;
            }
        }

        public static string Time_ConvertTimeFormat(string input, string format)
        {
            DateTimeFormatInfo dtFormat = new DateTimeFormatInfo();
            dtFormat.ShortDatePattern = format;

            DateTime dt = DateTime.Parse(input, dtFormat);

            return dt.ToString(s_dtFormat);
        }
    }
}
