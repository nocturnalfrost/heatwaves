using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace HeatWaves
{
    class Program
    {
        static void Main()
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

            // Load XML File and create Dict
            XElement catalog = XElement.Load(xmlPath);
            Dictionary<string, Dictionary<string, string>> bookDictionary = new Dictionary<string, Dictionary<string, string>>();

            // Parse XML Elements
            foreach (XElement book in catalog.Elements("book"))
            {
                string bookId = book.Attribute("id").Value;
                Dictionary<string, string> bookAttributes = new Dictionary<string, string>();
                ParseXml(book, bookAttributes, bookId);
                bookDictionary[bookId] = bookAttributes;
            }

            // Replace Placeholder in Word Doc
            foreach (var bookDict in bookDictionary)
            {
                ReplacePlaceholdersInWord(outputPath, bookDict.Value);
            }

            // For Debugging - Print Dict Elements to Console
            foreach (var book in bookDictionary)
            {
                Console.WriteLine($"Book ID: {book.Key}");
                foreach (var detail in book.Value)
                {
                    Console.WriteLine($"  {detail.Key}: {detail.Value}");
                }
            }
        }

        static void ParseXml(XElement element, Dictionary<string, string> dict, string parentKey = "")
        {
            string currentKey = string.IsNullOrEmpty(parentKey) ? element.Name.LocalName : $"{parentKey}.{element.Name.LocalName}";

            if (!element.HasElements)
            {
                dict[currentKey] = element.Value.Trim();
            }
            else
            {
                foreach (XElement child in element.Elements())
                {
                    ParseXml(child, dict, currentKey);
                }
            }
        }

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