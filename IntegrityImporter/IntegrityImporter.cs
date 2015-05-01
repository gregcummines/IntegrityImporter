using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using HtmlAgilityPack;
using log4net;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

namespace IntegrityImporter
{
    /// <summary>
    /// Describes what information we are currently looking at for parsing
    /// </summary>
    public enum ParserState
    {
        NothingInteresting,
        Requirement,
        Folder
    }

    public class IntegrityImporter
    {
        private static readonly ILog logger = LogManager.GetLogger(typeof(IntegrityImporter));

        public static List<Folder> ParseDocumentAsHtml(string filePath)
        {
            logger.Info("Entering ParseDocumentAsHtml...");
            // First save the word document as an Html file

            logger.Debug("Saving " + filePath + " as HTML");
            string htmlFilePath = SaveDocumentAsHtml(filePath);

            Dictionary<string, List<HtmlNode>> requirements = ParseRequirementsAsHtml(htmlFilePath);

            //if (requirements.Count() != srsDocument.Requirements.Count)
            //    logger.Error("Specification: " + srsDocument.Title + "Requirements count: Html=" + requirements.Count() + " WordMl=" + srsDocument.Requirements.Count);

            //foreach (Requirement requirement in srsDocument.Requirements)
            //{
            //    if (requirements.ContainsKey(requirement.RequirementNumber))
            //    {
            //        foreach (HtmlNode htmlNode in requirements[requirement.RequirementNumber])
            //        {
            //            if (htmlNode.OuterHtml.Length > 0)
            //            {
            //                requirement.Html.Add(htmlNode.OuterHtml);
            //            }
            //        }
            //    }
            //    else
            //    {
            //        logger.Warn("Requirement found by WordMl parser not found by Html parser: " + requirement.RequirementNumber);
            //    }
            //}

            //return srsDocument;
            return null;
        }

        private static Dictionary<string, List<HtmlNode>> ParseRequirementsAsHtml(string filePath)
        {
            // At this point filePath is now a HTML file...

            int req = 0;

            ParserState parserState = ParserState.NothingInteresting;
            var document = new HtmlDocument();
            document.OptionCheckSyntax = false;
            document.OptionFixNestedTags = false;
            document.OptionOutputOriginalCase = true;
            document.OptionWriteEmptyNodes = false;
            document.Load(filePath);

            Dictionary<string, List<HtmlNode>> requirements = new Dictionary<string, List<HtmlNode>>();
            List<HtmlNode> nodes = null;
            if (document.ParseErrors != null && document.ParseErrors.Count() > 0)
            {
                StringBuilder errorsString = new StringBuilder();
                foreach (var parseError in document.ParseErrors)
                {
                    errorsString.AppendLine(parseError.Reason);
                }
                Debug.WriteLine(errorsString.ToString());
            }
            else
            {
                int folderNestLevel = 0;

                int sectionMax = 10;
                for (int section = 1; section < sectionMax; section++)
                {
                    string xPathQuery = string.Format("//div[@class='WordSection{0}']", section);
                    HtmlNodeCollection wordSectionNode = document.DocumentNode.SelectNodes(xPathQuery);
                    if (wordSectionNode != null)
                    {

                        foreach (HtmlNode node in wordSectionNode[0].ChildNodes)
                        {
                            //// If we are not looking at requirement content, try to find a requirement
                            //if (parserState == ParserState.NothingInteresting)
                            //{
                            //    // If this is a paragraph element and our regular expression matches a requirement
                            //    if ((node.Name == "p") && (Regex.IsMatch(node.InnerText, regExRequirementSearch)))
                            //    {
                            //        // Add it to the dictionary
                            //        nodes = new List<HtmlNode>();
                            //        AddHtmlRequirement(filePath, node, regExRequirementNumber, nodes, requirements);
                            //        req++;
                            //        parserState = ParserState.Requirement;
                            //    }
                            //}
                            //else if (parserState == ParserState.Requirement)
                            //{
                            //    // If we find another requirement, add it as a new requirement to the dictionary
                            //    if ((node.Name == "p") && (Regex.IsMatch(node.InnerText, regExRequirementSearch)))
                            //    {
                            //        nodes = new List<HtmlNode>();
                            //        AddHtmlRequirement(filePath, node, regExRequirementNumber, nodes, requirements);
                            //        req++;
                            //        parserState = ParserState.Requirement;
                            //    }
                            //    // we found the requirements completion marker table, so nothing interesting
                            //    else if ((node.Name == "table") && (node.InnerText.Contains("Completed?")))
                            //    {
                            //        parserState = ParserState.NothingInteresting;
                            //    }
                            //    // keep adding content until either another requirement is found or we
                            //    // find the requirement completion marker table
                            //    else if (parserState == ParserState.Requirement)
                            //    {
                            //        UpdateImageNodes(filePath, node);
                            //        if (nodes != null)
                            //        {
                            //            nodes.Add(node);
                            //        }
                            //    }
                            //}
                        }
                    }
                }
            }

            Debug.WriteLine("Found " + req.ToString() + " requirements");
            return requirements;
        }



        /// <summary>
        /// Convert the img node that has a file reference into a data URI scheme
        /// <img src="data:image/jpg;base64,@(Html.Raw(Convert.ToBase64String((byte[])ViewBag.Image)))" alt="" /> 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="imageNode"></param>
        /// <returns></returns>
        private static void BuildBase64ImageNode(string filePath, HtmlNode imageNode)
        {
            string picFilePath = Path.GetDirectoryName(filePath);

            string source = imageNode.Attributes["src"].Value;
            string width = imageNode.Attributes["width"].Value;
            string height = imageNode.Attributes["height"].Value;

            source = source.Replace("%20", " ");
            picFilePath = Path.Combine(picFilePath, source);

            string ext = Path.GetExtension(picFilePath);

            byte[] fileData = null;
            using (
                BinaryReader binData =
                    new BinaryReader(File.Open(picFilePath, FileMode.Open)))
            {
                int length = (int)binData.BaseStream.Length;
                fileData = binData.ReadBytes(length);
                binData.Close();
            }
            string base64data = Convert.ToBase64String(fileData);

            imageNode.Attributes["src"].Value = "data:image/" + ext + ";base64," + base64data;
        }

        private static void AddHtmlRequirement(string filePath, HtmlNode node, string regExRequirementNumber, List<HtmlNode> nodes, Dictionary<string, List<HtmlNode>> requirements)
        {
            Match match = Regex.Match(node.InnerText, regExRequirementNumber);
            Debug.WriteLine("Requirement " + match.Value + " found");

            nodes.Add(node);
            if (!requirements.ContainsKey(match.Value))
                requirements.Add(match.Value, nodes);
            else
            {
                Debug.WriteLine("Duplicate requirement found " + match.Value + " in " + filePath);
            }
        }

        /// <summary>
        /// Updates all image nodes to use URI encoding
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="node"></param>
        private static void UpdateImageNodes(string filePath, HtmlNode node)
        {
            if (node != null)
            {
                if (node.Name == "img")
                    BuildBase64ImageNode(filePath, node);
                foreach (HtmlNode childNode in node.ChildNodes)
                {
                    UpdateImageNodes(filePath, childNode);
                }
            }
        }

        /// <summary>
        /// Saves a word document in Html format
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static string SaveDocumentAsHtml(string filePath)
        {
            Microsoft.Office.Interop.Word.Application oWordApplic = new Application();			// a reference to Word application 
            Microsoft.Office.Interop.Word.Document oDoc = null;                                  // a reference to the document 

            oWordApplic.Visible = false;

            object readOnly = true;
            object isVisible = false;
            object missing = System.Reflection.Missing.Value;
            string newFilePath = string.Empty;

            try
            {
                //Open this file
                object fileNameByRef = filePath;
                oDoc = oWordApplic.Documents.Open(ref fileNameByRef, ref missing, ref readOnly,
                                                  ref missing, ref missing, ref missing, ref missing, ref missing,
                                                  ref missing,
                                                  ref missing, ref missing, ref isVisible, ref missing, ref missing,
                                                  ref missing, ref missing);
                //oWordApplic.Visible = true;
                oDoc.ActiveWindow.View.RevisionsView = WdRevisionsView.wdRevisionsViewOriginal;
                //Show file
                //oDoc.Activate();


                //Unprotect so our find operation will work later.  This will throw an exception if the doc
                //is already Unprotected.  Just eat the exception.
                try
                {
                    oDoc.Unprotect(ref missing);
                }
                catch
                {
                }

                oDoc.ShowRevisions = true;

                //try
                //{
                //    oDoc.AcceptAllRevisions();
                //}
                //catch
                //{
                //}

                string path = Path.GetDirectoryName(filePath);
                path = Path.Combine(path, "html");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string newFileName = Path.GetFileNameWithoutExtension(filePath);
                newFileName += ".htm";
                newFilePath = Path.Combine(path, newFileName);

                oDoc.WebOptions.Encoding = MsoEncoding.msoEncodingUTF8;
                oDoc.SaveAs(newFilePath, WdSaveFormat.wdFormatFilteredHTML);
                oDoc.Close();
                oWordApplic.Quit();
                oDoc = null;

                var htmlSource = File.ReadAllText(newFilePath);
                var result = PreMailer.Net.PreMailer.MoveCssInline(htmlSource);
                File.WriteAllText(newFilePath, result.Html, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to convert {0} to html", filePath);
                throw;
            }
            finally
            {
                if (oDoc != null)
                {
                    oDoc.Close();
                    oWordApplic.Quit();
                    oDoc = null;
                }
            }
            return newFilePath;
        }


    }
}
