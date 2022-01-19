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
        public static Excel.Application oExcelApp;
        public static Excel.Workbook oLibroExcel; //Es como el documento tipo excel
        public static Excel.Worksheets oHojasExcel;
        public static Excel._Worksheet oHojaExcel;
        public static Excel.Range oRango;

        /// <summary>
        /// Encuentra el excel que esté abierto y crea los objetos con el documento activo
        /// </summary>
        public static void GetCurrentOpenExcel()
        {   
            oExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            oExcelApp.Visible = false;
            oLibroExcel = (Excel.Workbook)oExcelApp.ActiveWorkbook;
            oHojaExcel = (Excel._Worksheet)oLibroExcel.ActiveSheet;  
        }

        /// <summary>
        /// Busca un determinado caso de carga y borra todas las fila donde esté 
        /// </summary>
        /// <param name="oHojaExcel"></param>
        /// <param name="loadcase">Nombre del caso de carga a borrar</param>
        public static void DeleteLoadCases(Excel._Worksheet oHojaExcel, string loadcase)
        {
            int[] celdasMaximas = cFormatExcel.UsedRange(oHojaExcel);
            int columnaOutputCase = oHojaExcel.get_Range("A2", cFormatExcel.NumberToLL(celdasMaximas[1]) + "2").Find("Output Case").Column;
            string col = cFormatExcel.NumberToLL(columnaOutputCase);
            oRango = oHojaExcel.get_Range(col + "2", col + celdasMaximas[0].ToString());
            var valModal = oRango.Find(loadcase);
            while (valModal != null)
            {
                valModal = oRango.Find(loadcase);
                if (valModal != null) valModal.EntireRow.Delete();
            }
        }

        /// <summary>
        /// Borra la columna indicada
        /// </summary>
        /// <param name="oHojaExcel"></param>
        /// <param name="ColToDelete">Título de la columna en la fila 2 del excel generado por ETABS</param>
        public static void DeleteColumns(Excel._Worksheet oHojaExcel, string ColToDelete)
        {
            int[] celdasMaximas = cFormatExcel.UsedRange(oHojaExcel);
            var posColToDelete=oHojaExcel.get_Range("A2", cFormatExcel.NumberToLL(celdasMaximas[1]) + "2").Find(ColToDelete);
            if (posColToDelete != null) posColToDelete.EntireColumn.Delete();
        }

    }
}
