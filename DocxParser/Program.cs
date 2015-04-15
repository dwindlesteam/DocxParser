using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace DocxParser
{
    class Program
    {
        static void Main(string[] args)
        {
            StringCollection filePaths = new StringCollection();
            Console.WriteLine("Enter in the path to your Docx Folder. Make sure to escape with '\\'.");
            string folderPath = Console.ReadLine();

            try
            {
                filePaths = GetListOfPaths(folderPath.ToString());
            }
            catch (Exception e)
            {

                Console.WriteLine(e.ToString());
            }


            if (!(filePaths.Count > 0))
            {
                Console.WriteLine("No files found in DocxFolder.");
                Console.ReadLine();
            }

            else
            {
                foreach (string filePath in filePaths)
                {
                    try
                    {
                        Console.WriteLine(GetReportAndApproverInfo(filePath));
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }

                }
                Console.WriteLine("Done.");
                Console.ReadLine();
            }
        }

        /// <summary>
        /// Scans Docx folder for files and dumps the pathnames into StringCollection object
        /// </summary>
        /// <param name="folderPath"></param>
        /// <returns></returns>
        public static StringCollection GetListOfPaths(string folderPath)
        {
            StringCollection sc = new StringCollection();

            // string[] files = Directory.GetFiles(folderPath);
            var files = Directory.GetFiles(folderPath, "*.docx", SearchOption.AllDirectories);
     
            foreach (string fileName in files)
            {
                sc.Add(fileName);
            }
            return sc;
        }

        /// <summary>
        /// Parses .docx file and prints out the report name and approver information.
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public static string GetReportAndApproverInfo(string Path)
        {
            const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            StringBuilder textBuilder = new StringBuilder();

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(Path, false))
            {
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);

                XmlDocument xdoc = new XmlDocument(nt);
                xdoc.Load(wdDoc.MainDocumentPart.GetStream());

                XmlNodeList paragraphNodes = xdoc.SelectNodes("//w:tc", nsManager);

                foreach (XmlNode textNode in paragraphNodes)
                {
                    if (textNode.InnerText.Contains("Report Name"))
                    {
                        textBuilder.Append(textNode.InnerText);
                        textBuilder.Append(textNode.NextSibling.InnerText);
                        textBuilder.Append(Environment.NewLine);
                    }

                    if (textNode.InnerText.Contains("Approved By"))
                    {
                        textBuilder.Append(textNode.InnerText);
                        textBuilder.Append(textNode.NextSibling.InnerText);
                        textBuilder.Append(Environment.NewLine);
                        break;
                    }
                }
                wdDoc.Close();
            }
            return textBuilder.ToString();
        }
    }
}
