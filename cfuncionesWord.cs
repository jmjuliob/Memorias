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

        public static Word._Application oWordApp;
        public static Word._Document oDocumentoWord;

        /// <summary>
        /// Abre un nuevo Word y carga la plantilla de la empresa. (Primero finaliza cualquier Word.Process)
        /// </summary>
        public static void OpenWordTemplate()
        {           
            //Se cierran todos los procesos de word que esten abiertos
            System.Diagnostics.Process[] wordProcs = System.Diagnostics.Process.GetProcessesByName("WINWORD");

            foreach (System.Diagnostics.Process proc in wordProcs)
            {
                proc.Kill();
            }

            //Start Word 
            oWordApp = new Word.Application();
            oWordApp.Visible = false;
            //Crea un nuevo documento de word a partir de la plantilla selleccionada.
            object oTemplate = @"C:\Users\jmjul\Desktop\EFEPRIMACE\Programas\Word\FC.docx";
            oDocumentoWord = oWordApp.Documents.Add(ref oTemplate, ref oMissing, ref oMissing, ref oMissing);
        }


        
    }
}
