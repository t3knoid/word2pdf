using Application = Microsoft.Office.Interop.Word.Application;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace word2pdf
{
    class PDFConverter
    {
        public string wordFile {get; set;}

        public PDFConverter(string file)
        {
            this.wordFile = file;
        }

        public PDFConverter()
        {
        }

        public void Convert()
        {
            Application word = new Application();
            // Configure Word settings
            word.Visible = false;
            word.ScreenUpdating = false;
            object _MissingValue = Missing.Value;
            Document wordDoc = null;
            object filename = (object)wordFile;
            try
            {
                wordDoc = word.Documents.Open(ref filename, ref _MissingValue,
                     ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue,
                     ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue,
                     ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue);
                wordDoc.Activate();
                string pdfFileName = Path.ChangeExtension(wordFile, "pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;
                wordDoc.ExportAsFixedFormat(pdfFileName, WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("{0}",ex.Message));
            }
            finally
            {
                if (wordDoc !=null)
                {
                    wordDoc.Close(ref _MissingValue, ref _MissingValue, ref _MissingValue);
                }
               
            }
        }

    }
}
