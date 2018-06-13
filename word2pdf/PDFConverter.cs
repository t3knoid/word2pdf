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
                Console.WriteLine(ex.StackTrace);
                return;
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

                System.Threading.Thread.Sleep(500);
                Console.WriteLine("Waiting for Word to initialize");

                if (word.Documents.Count < 1)
                {
                    Console.WriteLine("Unable to open the specified document. If this is being executed by a service in a 64-bit environment, perform the following: \n(1) Launch C:\\Windows\\System32\\compexp.msc.\n(2) Open DCOM Config\\Microsoft Word 97 - 2003 Document Properties.\n(3) Set the Identity to 'The interactive user.'");
                    return;
                }

                try 
                {
                    wordDoc.Activate();
                    try
                    {
                        wordDoc.Fields.Update();
                        try
                        {
                            string pdfFileName = Path.ChangeExtension(WordFile, "pdf");
                            object fileFormat = WdSaveFormat.wdFormatPDF;
                            wordDoc.ExportAsFixedFormat(pdfFileName, WdExportFormat.wdExportFormatPDF);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(String.Format("Failed creating a PDF document. {0}", ex.Message));
                            Console.WriteLine(ex.StackTrace);
                            Console.WriteLine(ex.InnerException);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(String.Format("Failed to update fields. {0}", ex.Message));
                        Console.WriteLine(ex.StackTrace);
                        Console.WriteLine(ex.InnerException);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(String.Format("Failed to focus on document. {0}", ex.Message));
                    Console.WriteLine(ex.StackTrace);
                    Console.WriteLine(ex.InnerException);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(String.Format("Failed to open Word doc. {0}", ex.Message));
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.InnerException);
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
