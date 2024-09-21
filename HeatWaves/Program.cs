using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace HeatWaves
{
    internal class Program
    {
        private static void Main()
        {
            // Declare Variables
            string xmlPath = "C:\\Users\\mmhau\\Downloads\\Books.xml";
            string templatePath = "C:\\Users\\mmhau\\Downloads\\Books.docx";
            string currentDateTime = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string fileExtension = "docx";
            string outputPath = string.Concat("C:\\Users\\mmhau\\Downloads\\Books",
                                              currentDateTime,
                                              ".",
                                              fileExtension);

            // Create Copy of Template File
            File.Copy(templatePath, outputPath, true);

            // Load XML File
            XElement xmlData = XElement.Load(xmlPath);
            Dictionary<string, string> xmlDict = new();

            /*
            // For Debugging Reasons (XML Parsing)
            foreach (var keyValuePair in xmlDict)
            {
                Console.WriteLine($"{keyValuePair.Key}: {keyValuePair.Value}");
            }
            */

            // Iterate the XML File and save each XML key with its value in Dictionary
            ParseElement(xmlData, xmlDict);

            static void ParseElement(XElement element, Dictionary<string, string> dict)
            {
                if (!element.HasElements)
                {
                    dict[element.Name.LocalName] = element.Value.Trim();
                }
                else
                {
                    foreach (XElement child in element.Elements())
                    {
                        ParseElement(child, dict);
                    }
                }
            }

            // Replace Placeholder in Word Document
            ReplacePlaceholdersInWord(outputPath, xmlDict);

            static void ReplacePlaceholdersInWord(string docPath, Dictionary<string, string> xmlDict)
            {
                using WordprocessingDocument wordDoc = WordprocessingDocument.Open(docPath, true);
                var body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {
                    foreach (var key in xmlDict.Keys)
                    {
                        if (text.Text.Contains(key))
                        {
                            text.Text = text.Text.Replace(key, xmlDict[key]);
                        }
                    }
                }
            }
        }
    }
}