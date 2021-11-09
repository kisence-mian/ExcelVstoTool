using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula.Script
{
    public class ExcelFormula
    {
        #region Excel

        [ExcelFunction(Name = "Excel_Int2ColumnName", Description = "将数字转换为列的英文字母", Category = "Excel")]
        public static string GetColumnName(
            [ExcelArgument(Name = "Column", Description = "列的序号")]int index)
        {
            string columnName = "";

            while (index > 0)
            {
                var modulo = (index - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                index = (index - modulo) / 26;
            }

            return columnName;
        }

        [ExcelFunction(Name = "Excel_ColumnName2Int", Description = "将英文字母转换为数字", Category = "Excel")]
        public static int GetColumnIndex(
           [ExcelArgument(Name = "ColumnName", Description = "列的英文字母")]string name)
        {
            int result = 0;

            // A = 1
            // AA = 27
            // BA = 53
            // AAA = 
            // Z = 26

            for (int i = 0; i < name.Length; i++)
            {
                result += (Convert.ToInt32(name[i]) - 64) * (int)Math.Pow(26, name.Length - i - 1);
            }

            return result;
        }



        #endregion
    }
}
