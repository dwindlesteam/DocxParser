using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace DocxParser.Parser
{
    public class OpenXMLParser
    {
        private int _countOfApprovals;

        private int _countOfNonApprovals;

        public int CountOfApprovals 
        { 
            get 
            { 
                return _countOfApprovals; 
            } 
            private set
            {
                _countOfApprovals = value;
            }
        }

        public int CountOfNonApprovals 
        { 
            get 
            { 
                return _countOfNonApprovals; 
            }
            private set
            {
                _countOfNonApprovals = value;
            }
        }

        /// <summary>
        /// Parses .docx file and prints out the report name and approver information.
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public string GetReportAndApproverInfo(string Path)
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

                var reportName = paragraphNodes.Cast<XmlNode>()
                                               .Where(x => x.InnerText.Contains("Report Name"))
                                               .FirstOrDefault();

                if (reportName != null)
                {
                    textBuilder.Append(reportName.InnerText);
                    textBuilder.Append(reportName.NextSibling.InnerText);
                    textBuilder.Append(Environment.NewLine);
                }

                else
                {
                    textBuilder.Append("No Report Name");
                    textBuilder.Append(Environment.NewLine);
                }

                var approver = paragraphNodes.Cast<XmlNode>()
                                             .Where(x => x.InnerText.Contains("Approved By"))
                                             .FirstOrDefault();

                if (approver != null)
                {
                    textBuilder.Append(approver.InnerText);
                    textBuilder.Append(approver.NextSibling.InnerText);
                    textBuilder.Append(Environment.NewLine);
                    this.CountOfApprovals++;
                }

                else
                {
                    textBuilder.Append("No Approver");
                    textBuilder.Append(Environment.NewLine);
                    this.CountOfNonApprovals++;
                }

                wdDoc.Close();
            }

            return textBuilder.ToString();
        }
    }
}
