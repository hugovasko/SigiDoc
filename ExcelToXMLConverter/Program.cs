using OfficeOpenXml;
using System.Xml.Linq;

namespace ExcelToXMLConverter
{
    internal class Program
    {
        private static void Main()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                XDocument xmlTemplate;
                using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                {
                    xmlTemplate = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                }

                XDocument sealsList;
                using (var stream = new StreamReader(@"../all_seals.xml"))
                {
                    sealsList = XDocument.Load(stream);
                }

                XNamespace ns = "http://www.tei-c.org/ns/1.0";

                ExcelPackage package;
                using (var stream = new FileStream(@"./resources/newTest.xlsx", FileMode.Open, FileAccess.Read))
                {
                    package = new ExcelPackage(stream);
                }
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet == null)
                {
                    Console.WriteLine("The worksheet does not exist.");
                    return;
                }

                var dimensions = worksheet.Dimension;
                var headers = new Dictionary<string, int>();

                for (var row = 1; row <= dimensions.End.Row; row++)
                {
                    var header = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                    if (header == null) continue;
                    headers.Add(header, row);
                }

                for (var col = 2; col <= dimensions.End.Column; col++)
                {
                    var allValues = new Dictionary<string, string>();

                    foreach (var header in headers)
                    {
                        var value = worksheet.Cells[header.Value, col].Value?.ToString()?.Trim() ?? "―";
                        allValues.Add(header.Key, value);
                    }

                    var sealId = allValues["SIGIDOC ID"];
                    var filename = allValues["FILENAME"];
                    var sequence = sealId.PadLeft(4, '0');
                    allValues.Add("SEQUENCE", sequence);
                    allValues.Add("{}", "―");

                    var interpretiveText = allValues.ContainsKey("EDITION INTERPRETIVE") ? allValues["EDITION INTERPRETIVE"] : null;

                    if (string.IsNullOrEmpty(interpretiveText))
                    {
                        Console.WriteLine("Interpretive text is missing or empty.");
                        continue;
                    }

                    var lines = interpretiveText.Split('/');
                    var editionElement = xmlTemplate.Descendants(ns + "div").FirstOrDefault(e => (string)e.Attribute("type") == "edition" && (string)e.Attribute("subtype") == "editorial");

                    if (editionElement == null)
                    {
                        Console.WriteLine("Could not find the expected 'edition' div in the XML template.");
                        continue;
                    }

                    var textPartElement = editionElement.Descendants().FirstOrDefault(e => e.Name == ns + "div" && (string)e.Attribute("type") == "textpart" && (string)e.Attribute("n") == "obv");

                    if (textPartElement == null)
                    {
                        Console.WriteLine("Could not find the expected 'textpart' div with n='obv' in the XML template.");
                        continue;
                    }

                    textPartElement.Elements(ns + "ab").Remove();
                    string primaryIndentation = "    ";
                    string secondaryIndentation = primaryIndentation + "    ";

                    foreach (var line in lines)
                    {
                        var abElement = new XElement(ns + "ab", line.Trim());

                        textPartElement.Add(Environment.NewLine + secondaryIndentation);
                        textPartElement.Add(abElement);
                    }
                    textPartElement.Add(Environment.NewLine + primaryIndentation);

                    foreach (var element in xmlTemplate.Descendants())
                    {
                        foreach (var replacement in ReplacementValues.Replacements)
                        {
                            if (element.Value == replacement.key)
                            {
                                element.Value = allValues[replacement.value];
                            }

                            foreach (var attribute in element.Attributes())
                            {
                                if (attribute.Value == replacement.key)
                                {
                                    attribute.Value = allValues[replacement.value];
                                }
                            }
                        }
                    }

                    xmlTemplate.Save($"../webapps/ROOT/content/xml/epidoc/{filename}.xml");
                    allValues.Clear();

                    using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                    {
                        xmlTemplate = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                    }

                    var listElement = sealsList.Descendants(ns + "list").FirstOrDefault();

                    if (listElement != null)
                    {
                        bool isItemInList = listElement.Descendants(ns + "item").Any(i => i.Attribute("n")?.Value == filename);

                        if (!isItemInList)
                        {
                            var newListItem = new XElement(ns + "item");
                            newListItem.SetAttributeValue("n", filename);
                            newListItem.SetAttributeValue("sortKey", sequence);
                            listElement.Add(newListItem);

                            var sortedListItems = listElement.Descendants(ns + "item").OrderBy(i => i.Attribute("sortKey")?.Value).ToList();

                            listElement.ReplaceNodes(sortedListItems);

                            sealsList.Save(@"../all_seals.xml");

                            sealsList = XDocument.Load(@"../all_seals.xml");
                        }
                    }
                }
                package.Dispose();
                Console.WriteLine("Success!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}