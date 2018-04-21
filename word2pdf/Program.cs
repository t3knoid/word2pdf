using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace word2pdf
{
    class Program
    {
        /// <summary>
        /// word2pdf converts a single word document to PDF. It can also convert a
        /// set of documents by specifying a folder containing files with doc and docx 
        /// extensions
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var parser = new CommandLine();
            parser.Parse(args);

            if (parser.Arguments.Count > 0)
            {
                if (parser.Arguments.ContainsKey("convert"))
                {
                    WordFiles wordfiles = new WordFiles(parser.Arguments["convert"][0]);
                    if (wordfiles.Count > 0)
                    {
                        foreach (string wordfile in wordfiles.Docfiles)
                        {
                            PDFConverter pdfConverter = new PDFConverter();
                            pdfConverter.WordFile = @wordfile;
                            pdfConverter.Convert();
                        }
                    }
                    else
                    {
                        Console.WriteLine("Nothing to do.");
                    }
                }
                else
                {
                    usage();
                }
            }
            else
            {
                usage();
            }
        }
        static void usage()
        {
            Console.WriteLine("word2pdf -convert file [folder]");
        }
    }
}
