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
        static void Main(string[] args)
        {
            var parser = new CommandLine();
            parser.Parse(args);

            if (parser.Arguments.Count > 0)
            {
                WordFiles wordfiles = new WordFiles(parser.Arguments["convert"][0]);

                if (parser.Arguments.ContainsKey("convert"))
                {
                    PDFConverter pdfConverter = new PDFConverter();
                    pdfConverter.wordFile = @parser.Arguments["convert"][0];
                    pdfConverter.Convert();
                }
                else
                {
                    Console.WriteLine("Nothing to do.");
                }
            }
            else
            {
                Console.WriteLine("Nothing to do.");
            }
        }
    }
}
