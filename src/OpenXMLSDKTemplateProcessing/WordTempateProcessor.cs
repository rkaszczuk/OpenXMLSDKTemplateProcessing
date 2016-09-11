using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OpenXMLSDKTemplateProcessing
{
    public class WordTempateProcessor
    {
        public static string ReadAllFromDocument(string path)
        {
            string documentText = String.Empty;
            using (var wordDocument = WordprocessingDocument.Open(path, false))
            {
                foreach (var paragraph in wordDocument.MainDocumentPart.Document.Descendants<Paragraph>())
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        documentText += run.InnerText;
                    }

                }
            }
            return documentText;
        }
        public static byte[] CreateDocumentFromTemplate(string templatePath, Dictionary<string, string> tagValueDictionary)
        {
            byte[] templateBytes = File.ReadAllBytes(templatePath);


            using (var templateStream = new MemoryStream())
            {

                templateStream.Write(templateBytes, 0, templateBytes.Length);
                using (var wordDoc = WordprocessingDocument.Open(templateStream, true))
                {
                    foreach (var par in wordDoc.MainDocumentPart.Document.Descendants<Paragraph>())
                    {
                        foreach (var tagValue in tagValueDictionary)
                        {
                            ReplaceTextInParagraph(par, tagValue.Key, tagValue.Value);
                        }
                    }
                }

                templateStream.Position = 0;
                var result = templateStream.ToArray();
                templateStream.Flush();
                return result;
            }
        }
        private static void ReplaceTextInParagraph(Paragraph paragraph, string oldValue, string newValue)
        {
            if (paragraph.InnerText.Contains(oldValue))
            {
                Run newRun = (paragraph.Descendants<Run>().FirstOrDefault().CloneNode(true) ?? new Run()) as Run;
                newRun.RemoveAllChildren<Text>();
                newRun.AppendChild(new Text(paragraph.InnerText.Replace(oldValue, newValue)));
                paragraph.RemoveAllChildren<Run>();
                paragraph.AppendChild(newRun);
            }
        }
    }
}
