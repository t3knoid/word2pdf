using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace word2pdf
{
    /// <summary>
    /// A list of Word documents enumerated from the specified folder.
    /// </summary>
    class WordFiles
    {
        public List<string> Docfiles = new List<string>();
        public List<string> Ext { get; set; }
        public string Param { get; set; }
        public int Count
        {
            get
            {
                return Docfiles.Count();
            }
        }

        public WordFiles(string p)
        {
            this.Param = Path.GetFullPath(p);
            GetFileList();
       }

        /// <summary>
        /// Enumerates the files from the given path
        /// </summary>
        private void GetFileList()
        {
            try
            {
                FileAttributes attr = File.GetAttributes(Param);
                if (attr.HasFlag(FileAttributes.Directory))
                {
                    foreach (string file in Directory.GetFiles(Param, "*.doc"))
                    {
                        Docfiles.Add(file);
                    }
                }
                else
                {
                    Docfiles.Add(Param);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
