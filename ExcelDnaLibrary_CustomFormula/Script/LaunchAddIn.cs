using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDnaLibrary_CustomFormula
{
    [ComVisible(false)]
    public class LaunchAddIn : IExcelAddIn
    {
        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }

        public void AutoOpen()
        {
            
            IntelliSenseServer.Install();
        }
    }
}
