using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelVstoTool.DialogWindow
{
    public class WindowUntilTool
    {
        public static bool CheckRangeFormat(string content)
        {
            return Regex.IsMatch(content, "^([\\s\\S]*)![A-Z]+[0-9]+:[A-Z]+[0-9]+$");
        }
    }
}
