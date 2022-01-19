using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


enum eTableStories
{
    Niveles, Altura, Elevación
}



namespace WFAAPIETABS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cFuncionesEtabs.Open_Etabs();
            double[] Periods = { };
            double[] UX = { }; double[] UY = { }; double[] UZ = { }; 
            double[] RX = { }; double[] RY = { }; double[] RZ = { };

            cFuncionesEtabs.Open_Etabs();
            cFuncionesEtabs.Get_ModalPeriods(ref Periods, ref UX, ref UY, ref UZ, ref RX, ref RY, ref RZ);

            string[] StoryNames = { }; double[] StoryElevations = { }; double[] StoryHeights = { };

            cFuncionesEtabs.ObtenerAluturaEdificio(ref StoryNames,ref StoryHeights,ref StoryElevations);

            string[] TableKey = { }; int NumberTables = 0;
            string[] TableName = { }; int[] ImportType = { }; bool[] IsEmpty = { };
            //cFuncionesEtabs.GetTablesKeys(ref TableKey, ref NumberTables);
            cFuncionesEtabs.GetTablesKeys(ref TableKey, ref NumberTables, ref TableName, ref ImportType, ref IsEmpty);

            

            string[] TablasExcel = { "Base Reactions", "Diaphragm Max Over Avg Drifts" };

            cFuncionesEtabs.TablesToExcel(ref TablasExcel);

            cfuncionesExcel.GetCurrentOpenExcel();

            /*
             * NOTA IMPORTANTE: no se recomienda usar foreach con las librerías interop 
             * porque pueden conducir a excepciones
             **/ 

            for (int i = 1; i <= TablasExcel.Length; i++)
            {
                cfuncionesExcel.oHojaExcel = (Excel._Worksheet)cfuncionesExcel.oLibroExcel.Worksheets[i];
                cfuncionesExcel.DeleteLoadCases(cfuncionesExcel.oHojaExcel,"Modal");
                cfuncionesExcel.DeleteColumns(cfuncionesExcel.oHojaExcel, "Step Type");
                cfuncionesExcel.DeleteColumns(cfuncionesExcel.oHojaExcel, "Step Number");
                cFormatExcel.DarFormato(cfuncionesExcel.oHojaExcel);

                cfuncionesExcel.oRango= cfuncionesExcel.oHojaExcel.UsedRange;
                cfuncionesExcel.oRango.Copy();
                if (i==1)
                {
                    cfuncionesWord.OpenWordTemplate();
                    Word.Paragraph oPara1 = cfuncionesWord.oDocumentoWord.Content.Paragraphs.Add(ref cfuncionesWord.oMissing);
                    oPara1.Range.Text = TablasExcel[i-1];
                    oPara1.Range.InsertParagraphAfter();
                    oPara1.Range.Paste();
                    cFormatWord.RepeatTitleRows(cfuncionesWord.oDocumentoWord.Content.Tables[i]);
                }
                else
                {
                    Word.Paragraph oPara1 = cfuncionesWord.oDocumentoWord.Content.Paragraphs.Add(ref cfuncionesWord.oMissing);
                    oPara1.Range.Text = TablasExcel[i - 1];
                    oPara1.Range.InsertParagraphAfter();
                    oPara1.Range.Paste();
                    cFormatWord.RepeatTitleRows(cfuncionesWord.oDocumentoWord.Content.Tables[i]);
                }
            }

            cfuncionesWord.oDocumentoWord.SaveAs2(@"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\test.docx");
            cfuncionesWord.oDocumentoWord.Close();

            cfuncionesExcel.oLibroExcel.SaveAs(@"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\test.xlsx");
            cfuncionesExcel.oLibroExcel.Close();




            //Excel.Range oRango = null;

            //try
            //{
            //    oExcelApp = new Excel.Application();
            //    oExcelApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            //    oLibroExcel = oExcelApp.ActiveWorkbook;
            //    oHojaExcel = (Excel._Worksheet)oLibroExcel.ActiveSheet;
            //    oHojaExcel.get_Range("A4", "C4").Font.Name = "Courier New";



            //    //oLibroExcel = (Excel._Workbook)(oExcelApp.Workbooks.Add(Missing.Value));

            //}
            //finally
            //{
            //    if (oLibroExcel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oLibroExcel);
            //    if (oHojaExcel != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oHojaExcel);

            //}


            //Excel.Application oExcelApp;

            //oExcelApp = new Excel.Application();





            //Excel._Workbook oLibroExcel;
            //Excel._Worksheet oHojaExcel;
            //Excel.Range oRango;

            ////Start Excel and get Application object.
            ////oExcelApp = new Excel.Application();


            ////Get a new workbook.
            ////oExcelApp.Visible = true;

            //oLibroExcel = oExcelApp.ActiveWorkbook;


            //oLibroExcel = (Excel._Workbook)(oExcelApp.Workbooks.Add(Missing.Value));


            //oHojaExcel = (Excel._Worksheet)oLibroExcel.ActiveSheet;

            //try
            //{
            //    //Add table headers going cell by cell.                
            //    for (int i = 1; i < Enum.GetNames(typeof(eTableStories)).Length+1 ; i++)
            //    {
            //        oHojaExcel.Cells[1, i] = ((eTableStories)i - 1).ToString();

            //        for (int j = 2; j < StoryNames.Count()+1; j++)
            //        {
            //            if (i==1)
            //            {
            //                oHojaExcel.Cells[j, i] = StoryNames[j - 2+1];
            //            }
            //            else if (i == 2)
            //            {
            //                oHojaExcel.Cells[j, i] = Math.Round(StoryElevations[j - 2 + 1], 2);
            //            }
            //            else if (i == 3)
            //            {
            //                oHojaExcel.Cells[j, i] = Math.Round(StoryHeights[j - 2 + 1], 2);
            //            }
            //        }
            //    }


            //    oHojaExcel.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Font.Name = "Courier New";
            //    oHojaExcel.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + "1").Font.Bold = true;
            //    oHojaExcel.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + "1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //    oHojaExcel.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Borders.Weight= Excel.XlBorderWeight.xlThin;


            //}
            //catch (Exception theException)
            //{
            //    String errorMessage;
            //    errorMessage = "Error: ";
            //    errorMessage = String.Concat(errorMessage, theException.Message);
            //    errorMessage = String.Concat(errorMessage, " Line: ");
            //    errorMessage = String.Concat(errorMessage, theException.Source);

            //    MessageBox.Show(errorMessage, "Error");
            //}


            //object oMissing = System.Reflection.Missing.Value;
            //object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            ////Start Word and create a new document.
            //Word._Application oWordApp;
            //Word._Document oDocumentoWord;
            //oWordApp = new Word.Application();
            //oWordApp.Visible = false;
            ////oDocumentoWord = oWordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            //object oTemplate = @"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\FC.docx";
            //oDocumentoWord = oWordApp.Documents.Add(ref oTemplate, ref oMissing,
            //ref oMissing, ref oMissing);

            //Word.Paragraph oPara1 = oDocumentoWord.Content.Paragraphs.Add(ref oMissing);
            //oPara1.Range.Text = "Heading 1";
            //oPara1.Range.Font.Bold = 1;
            //oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            //oPara1.Range.InsertParagraphAfter();

            //Word.Table firstTable = oDocumentoWord.Tables.Add(oPara1.Range, 1, 1, ref oMissing, ref oMissing);

            //oHojaExcel.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Copy();

            //firstTable.Rows.Select();

            //firstTable.Range.Paste();

            //oDocumentoWord.SaveAs2(@"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\test.docx");
            //oDocumentoWord.Close();

            //oLibroExcel.SaveAs(@"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\test.xlsx");
            //oLibroExcel.Close();


            //firstTable.Borders.Enable = 1;

            //foreach (Word.Row row in firstTable.Rows)
            //{
            //    foreach (Word.Cell cell in row.Cells)
            //    {
            //        //Header row  
            //        if (cell.RowIndex == 1)
            //        {
            //            cell.Range.Text = ((eTableTitles)cell.ColumnIndex).ToString();
            //            cell.Range.Font.Bold = 1;
            //            //other format properties goes here  
            //            cell.Range.Font.Name = "verdana";
            //            cell.Range.Font.Size = 11;
            //            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
            //            cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
            //            //Center alignment for the Header cells  
            //            cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            //            cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            //        }
            //        //Data row  
            //        else
            //        {
            //            cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
            //        }
            //    }
            //}





        }
    }
}