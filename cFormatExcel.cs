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
        public enum eNNtoAN
        {
            A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z
        }

        public static int[] UsedRange(Excel._Worksheet oSheet) //Encuentra el rango en el que hay data
        {
            oRng = oSheet.UsedRange;
            int numCols = oRng.Columns.Count;
            int numRows = oRng.Rows.Count;
            

            return new int[2] { numRows, numCols};
        }

        public static string NumberToLL(int col/*, int row*/) //Encuentra la columna expresada en formato "A1" like 1,1 => A,1
        {
            return ((eNNtoAN)col).ToString();
        }

        public static void TableFormat(Excel._Worksheet oSheet)
        {
            string FinalSelection = NumberToLL(UsedRange(oSheet)[1]) + UsedRange(oSheet)[0].ToString(); //Encuentra la última celda del rango a seleccionar para darle formato
            oRng = oSheet.get_Range("A1", FinalSelection);
            oRng.Font.Name = "Courier New";
            oRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            oRng.Borders.Weight = Excel.XlBorderWeight.xlThin;
            //oRng.NumberFormat = "0.00";
        }


        public static void TableHeadder(Excel._Worksheet oSheet)
        {
            string FinalSelectionH = NumberToLL(UsedRange(oSheet)[1]) + "1";
            oRng = oSheet.get_Range("A1", FinalSelectionH);
            oRng.MergeCells = true;
            oRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            oRng.Interior.TintAndShade = -0.249977111117893;
            oRng.Merge();

            string FinalSelection = NumberToLL(UsedRange(oSheet)[1]) + "3"; //Encuentra la última celda del rango a seleccionar para darle formato
            oRng = oSheet.get_Range("A1", FinalSelection);
            oRng.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
            oRng.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            oRng.Interior.TintAndShade = -0.149998474074526;
            
        }

        public static void TableCore(Excel._Worksheet oSheet)
        {
            string FinalSelection = NumberToLL(UsedRange(oSheet)[1]) + UsedRange(oSheet)[0].ToString(); //Encuentra la última celda del rango a seleccionar para darle formato
            oSheet.get_Range("A4", FinalSelection).BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium);
        }

        public static void DarFormato(Excel._Worksheet oSheet)
        {
            TableFormat(oSheet);

            TableHeadder(oSheet);

            TableCore(oSheet);
        }


        //oSheet.get_Range("A4", "C4").Font.Name = "Courier New";
        //oSheet.get_Range("A1", ((eNNtoAN) Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Borders.Weight= Excel.XlBorderWeight.xlThin;


    }
}
