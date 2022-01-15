using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace WFAAPIETABS
{
    public class cfuncionesExcel
    {   
        public static Excel.Application oXL;
        public static Excel.Workbook oWB;
        public static Excel.Worksheets oSheets;
        public static Excel._Worksheet oSheet;
        public static Excel.Range oRng;


        public static void GetCurrentOpenExcel()
        {
            
            
            oXL = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = false;
            oWB = (Excel.Workbook)oXL.ActiveWorkbook;
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;
            

        }
    }
}
