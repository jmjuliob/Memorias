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
    class cFormatWord : cfuncionesWord
    {
        public static void RepeatTitleRows(Word.Table table)
        {
            table.Rows[1].HeadingFormat = -1;
            table.Rows[2].HeadingFormat = -1;
            table.Rows[3].HeadingFormat = -1;
        }

    }
}
