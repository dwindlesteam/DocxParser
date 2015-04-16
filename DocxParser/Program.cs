using DocxParser.Helper;
using DocxParser.Parser;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace DocxParser
{
    class Program
    {
        static void Main(string[] args)
        {
            OpenXMLParser parser = new OpenXMLParser();
            DirectoryOperator directoryOperator = new DirectoryOperator();

            IList<string> filePaths = new List<string>();
            IList<string> reportNameAndApprovalList = new List<string>();

            Console.WriteLine(@"Enter the path to your Docx Folder. Make sure to escape with '\\'.");
            Console.WriteLine(@"Example: C:\\Users\\foo\\Desktop\\DocxFolder");
            string folderPath = Console.ReadLine();
            Console.WriteLine(@"Enter the path to your text file + textfile name. Make sure to escape with '\\'.");
            string textFilePath = Console.ReadLine();

            try
            {
                filePaths = directoryOperator.GetListOfAllPaths(folderPath.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            if (!(filePaths.Count > 0))
            {
                Console.WriteLine("No files found in {0}.", folderPath);
                Console.ReadLine();
            }

            else
            {
                Console.WriteLine("Processing...");
                foreach (string filePath in filePaths)
                {
                    try
                    {
                        reportNameAndApprovalList.Add(parser.GetReportAndApproverInfo(filePath));
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }

                try
                {
                    System.IO.File.WriteAllLines(textFilePath, reportNameAndApprovalList);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            
                Console.WriteLine("{0} files written to {1}", (parser.CountOfApprovals + parser.CountOfNonApprovals).ToString(), textFilePath);
                Console.WriteLine("# of Approvals: {0}", parser.CountOfApprovals.ToString());
                Console.WriteLine("# of NonApprovals: {0}", parser.CountOfNonApprovals.ToString());
                Console.WriteLine("Done.");
                Console.ReadLine();
            }
        }
    }
}
