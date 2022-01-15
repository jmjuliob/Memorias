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
    class cfuncionesWord
    {
        public static object oMissing = System.Reflection.Missing.Value;
        public static object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

        //Start Word and create a new document.
        public static Word._Application oWord;
        public static Word._Document oDoc;

        public static void OpenWordTemplate()
        {
            //Se cierran todos los procesos de word que esten abiertos
            System.Diagnostics.Process[] wordProcs = System.Diagnostics.Process.GetProcessesByName("WINWORD");
            //System.Diagnostics.Process[] wordProcs = System.Diagnostics.Process.GetProcesses();

            foreach (System.Diagnostics.Process proc in wordProcs)
            {
                proc.Kill();
            }


            oWord = new Word.Application();
            oWord.Visible = false;
            //oDoc = oWord.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            object oTemplate = @"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\FC.docx";
            oDoc = oWord.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);


        }


        
    }
}
