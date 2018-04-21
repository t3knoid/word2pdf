using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace word2pdf
{
    class WordFiles
    {
        public List<string> files = new List<string>();
        public List<string> ext { get; set; }
        public string path { get; set; }

        public WordFiles(string p)
        {
            this.ext = new List<string> {"doc","docx"};
            this.path = p;
            GetFileList();
       }

        /// <summary>
        /// Enumerates the files from the given path
        /// </summary>
        private void GetFileList()
        {
            try
            {
                FileAttributes attr = File.GetAttributes(@path);
                if (attr.HasFlag(FileAttributes.Directory))
                {
                    // Read files from the directory
                    foreach (string file in Directory.GetFiles(path, "*.*", SearchOption.AllDirectories).Where(s => ext.Contains(Path.GetExtension(s))))
                    {
                        files.Add(file);
                    }
                }
                else
                {
                    files.Add(path);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
