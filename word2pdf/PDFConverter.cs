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
    /// <summary>
    /// Converts a Word document to PDF format
    /// </summary>
    class PDFConverter
    {
        /// <summary>
        /// A fully qualified path to a word document.
        /// </summary>
        public string WordFile {get; set;}

        public PDFConverter(string file)
        {
            this.WordFile = file;
        }

        public PDFConverter()
        {
        }
        /// <summary>
        /// Converts a Word document to pdf. Specify the Word document using the
        /// wordFile property.
        /// </summary>
        public void Convert()
        {
            Application word = null;

            try
            {
                word = new Application
                {
                    // Configure Word settings
                    Visible = false,
                    ScreenUpdating = false
                };
            }
            catch (Exception ex)
            {
                Console.Write(String.Format("Failed to initialize Word. {0}",ex.Message));
                Environment.Exit(1);
            }

            object _MissingValue = Missing.Value;
            Document wordDoc = null;
            object filename = (object)WordFile;
            try
            {
                //wordDoc = word.Documents.Open(ref filename, Visible: true);
                wordDoc = word.Documents.Open(ref filename, ref _MissingValue,
                     ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue,
                     ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue,
                     ref _MissingValue, ref _MissingValue, ref _MissingValue, ref _MissingValue);
                wordDoc.Fields.Update();
            }
            catch (Exception ex){
                Console.WriteLine(String.Format("Failed to open Word doc. {0} ", ex.ToString()));
            }
            try
            {
                wordDoc.Activate();
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Failed to focus on document. {0} ", ex.ToString()));
            }
            try 
            {
                string pdfFileName = Path.ChangeExtension(WordFile, "pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;
                wordDoc.ExportAsFixedFormat(pdfFileName, WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Failed creating a PDF document. {0} ", ex.ToString()));
            }
            finally
            {
                if (wordDoc !=null)
                {
                    Console.WriteLine("Closing document.");
                    wordDoc.Close(WdSaveOptions.wdDoNotSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat, false);
                }
                if (word != null)
                {
                    Console.WriteLine("Exiting Word.");
                    word.Quit();
                }
               
            }
        }

    }
}
