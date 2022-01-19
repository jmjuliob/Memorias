using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace WFAAPIETABS
{
    class cFormatExcel:cfuncionesExcel
    {
        // Enum con la info para cambiar celdas de excel de formato 1,1 a formato "A,1"
        public enum eNNtoAN
        {
            A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
        }

        /// <summary>
        /// Método para encontrar el número de celdas y columnas que tienen data
        /// </summary>
        /// <param name="oHojaExcel">Objeto con la hoja activa de excel</param>
        /// <returns>Lista con [Número de filas, Número de columnas]</returns>
        public static int[] UsedRange(Excel._Worksheet oHojaExcel) //Encuentra el rango en el que hay data
        {
            oRango = oHojaExcel.UsedRange;
            int numCols = oRango.Columns.Count;
            int numRows = oRango.Rows.Count;
            

            return new int[2] { numRows, numCols};
        }

        /// <summary>
        /// Cambia la celda de excel de formato 1,1 a formato "A,1"
        /// </summary>
        /// <param name="col"></param>
        /// <returns>In 3, out C</returns>
        public static string NumberToLL(int col) 
        {
            return ((eNNtoAN)col).ToString();
        }

        /// <summary>
        /// Asigna tipo de letra, alineación y bordes
        /// </summary>
        /// <param name="oHojaExcel">Objeto con la hoja activa de excel</param>
        public static void TableFormat(Excel._Worksheet oHojaExcel)
        {
            string FinalSelection = NumberToLL(UsedRange(oHojaExcel)[1]) + UsedRange(oHojaExcel)[0].ToString(); //Encuentra la última celda del rango a seleccionar para darle formato
            oRango = oHojaExcel.get_Range("A1", FinalSelection);
            oRango.Font.Name = "Courier New";
            oRango.Borders.Weight = Excel.XlBorderWeight.xlThin;
            //oRango.NumberFormat = "0.00";
        }

        /// <summary>
        /// Asigna color de fondo a las celdas del encabezado, bordes exteriores gruesos y combina y centra el título de la tabla
        /// </summary>
        /// <param name="oHojaExcel">Objeto con la hoja activa de excel</param>
        public static void TableHeadder(Excel._Worksheet oHojaExcel)
        {
            string FinalSelectionH = NumberToLL(UsedRange(oHojaExcel)[1]) + "1"; //Encuentra la última celda del rango  del título de la tabla para darle formato
            oRango = oHojaExcel.get_Range("A1", FinalSelectionH);
            oRango.MergeCells = true;
            oRango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            oRango.Interior.TintAndShade = -0.249977111117893;
            oRango.Merge();

            FinalSelectionH = NumberToLL(UsedRange(oHojaExcel)[1]) + "2";
            oRango = oHojaExcel.get_Range("A2", FinalSelectionH);

            string FinalSelection = NumberToLL(UsedRange(oHojaExcel)[1]) + "3"; //Encuentra la última celda del resto del encabezado para obtener el rango a seleccionar y darle formato
            oRango = oHojaExcel.get_Range("A1", FinalSelection);
            oRango.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            oRango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            oRango.Interior.TintAndShade = -0.149998474074526;
            oRango.WrapText = true;
            oRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            oRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

        }

        /// <summary>
        /// Asigna bordes gruesos al cuerpo de la tabla
        /// </summary>
        /// <param name="oHojaExcel">Objeto con la hoja activa de excel</param>
        public static void TableCore(Excel._Worksheet oHojaExcel)
        {
            string FinalSelection = NumberToLL(UsedRange(oHojaExcel)[1]) + UsedRange(oHojaExcel)[0].ToString(); //Encuentra la última celda del rango a seleccionar para darle formato
            oHojaExcel.get_Range("A4", FinalSelection).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            //oHojaExcel.get_Range("A4", FinalSelection).Columns.AutoFit();
        }

        /// <summary>
        /// Le asigna el formato predefinido a la hoja de cálculo seleccionada
        /// </summary>
        /// <param name="oHojaExcel">Objeto con la "sheet</param>
        public static void DarFormato(Excel._Worksheet oHojaExcel)
        {
            TableFormat(oHojaExcel);

            TableHeadder(oHojaExcel);

            TableCore(oHojaExcel);
        }


        //oHojaExcel.get_Range("A4", "C4").Font.Name = "Courier New";
        //oHojaExcel.get_Range("A1", ((eNNtoAN) Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Borders.Weight= Excel.XlBorderWeight.xlThin;


    }
}
