using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Microsoft.Office.Interop.Word;
using System.IO;

namespace WordDocReplaceTest
{
    class Program
    {       
        static void Main(string[] args)
        {
            //using Word = Microsoft.Office.Interop.Word;

            object fileName = Path.Combine(@"D:\VS2010Workspace\WordDocReplaceTest\WordDocReplaceTest\bin\Release", "TestDoc.docx");
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };

            Microsoft.Office.Interop.Word.Document aDoc = wordApp.Documents.Open(ref fileName, ReadOnly: false, Visible: true);

            aDoc.Activate();

            Microsoft.Office.Interop.Word.Find fnd = wordApp.ActiveWindow.Selection.Find;

            fnd.ClearFormatting();
            fnd.Replacement.ClearFormatting();
            fnd.Forward = true;
            fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

            fnd.Text = "{替换前内容}";
            fnd.Replacement.Text = "替换后内容-updated";

            fnd.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
            aDoc.Save();

            aDoc.Close(ref missing, ref missing, ref missing);
            wordApp.Quit(ref missing, ref missing, ref missing);
        }
    }
}
