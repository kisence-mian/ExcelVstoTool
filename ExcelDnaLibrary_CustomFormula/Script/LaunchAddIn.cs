using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula
{
    public class LaunchAddIn : IExcelAddIn
    {
        public void AutoClose()
        {
            //IntelliSenseServer.Install();
        }

        public void AutoOpen()
        {
            //IntelliSenseServer.Uninstall();
        }
    }
}
