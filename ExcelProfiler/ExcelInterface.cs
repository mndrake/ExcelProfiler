namespace ExcelProfiler
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using ExcelDna.Integration;
    
    public class ExcelInterface : IExcelAddIn
    {

        public void AutoClose()
        {
            // insert any clean up steps needed when the add-in is unloaded
        }

        public void AutoOpen()
        {
            // insert any steps needed when the add-in is loaded
        }
    }
}