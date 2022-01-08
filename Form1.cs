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

enum eNNtoAN
{
    A=1,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z
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
            double[] UX = { }; double[] UY = { }; double[] UZ = { }; double[] RX = { };
            double[] RY = { }; double[] RZ = { };

            cFuncionesEtabs.Open_Etabs();
            cFuncionesEtabs.Get_ModalPeriods(ref Periods, ref UX, ref UY, ref UZ, ref RX, ref RY, ref RZ);

            string[] StoryNames = { }; double[] StoryElevations = { }; double[] StoryHeights = { };

            cFuncionesEtabs.ObtenerAluturaEdificio(ref StoryNames,ref StoryHeights,ref StoryElevations);

            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            //Start Excel and get Application object.
            oXL = new Excel.Application();
            oXL.Visible = false;

            //Get a new workbook.
            oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
            oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            try
            {
                //Add table headers going cell by cell.                
                for (int i = 1; i < Enum.GetNames(typeof(eTableStories)).Length+1 ; i++)
                {
                    oSheet.Cells[1, i] = ((eTableStories)i - 1).ToString();

                    for (int j = 2; j < StoryNames.Count()+1; j++)
                    {
                        if (i==1)
                        {
                            oSheet.Cells[j, i] = StoryNames[j - 2+1];
                        }
                        else if (i == 2)
                        {
                            oSheet.Cells[j, i] = Math.Round(StoryElevations[j - 2 + 1], 2);
                        }
                        else if (i == 3)
                        {
                            oSheet.Cells[j, i] = Math.Round(StoryHeights[j - 2 + 1], 2);
                        }
                    }
                }


                oSheet.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Font.Name = "Courier New";
                oSheet.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + "1").Font.Bold = true;
                oSheet.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + "1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                oSheet.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Borders.Weight= Excel.XlBorderWeight.xlThin;


            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }


            object oMissing = System.Reflection.Missing.Value;
            object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.
            Word._Application oWord;
            Word._Document oDoc;
            oWord = new Word.Application();
            oWord.Visible = false;
            //oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            object oTemplate = @"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\FC.docx";
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing,
            ref oMissing, ref oMissing);

            Word.Paragraph oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Text = "Heading 1";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();

            Word.Table firstTable = oDoc.Tables.Add(oPara1.Range, 1, 1, ref oMissing, ref oMissing);

            oSheet.get_Range("A1", ((eNNtoAN)Enum.GetNames(typeof(eTableStories)).Length).ToString() + StoryNames.Count().ToString()).Copy();

            firstTable.Rows.Select();
            
            firstTable.Range.Paste();

            oDoc.SaveAs2(@"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\test.docx");
            oDoc.Close();

            oWB.SaveAs(@"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\test.xlsx");
            oWB.Close();


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