using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;

namespace DocxParser.Helper
{
    public class DirectoryOperator
    {
        /// <summary>
        /// Scans Docx folder for files and dumps the pathnames into StringCollection object
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public IList<string> GetListOfAllPaths(string folderPath)
        {
           // StringCollection sc = new StringCollection();
            IList<string> filePathList = new List<string>();

            var files = Directory.GetFiles(folderPath, "*.docx", SearchOption.AllDirectories);

            foreach (string fileName in files)
            {
                filePathList.Add(fileName); 
            }

            return filePathList;
        }
    }
}
