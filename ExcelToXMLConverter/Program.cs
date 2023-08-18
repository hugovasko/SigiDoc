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
                using (var stream = new FileStream(@"./resources/test.xlsx", FileMode.Open, FileAccess.Read))
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

                    var sealId = allValues["SEAL ID (IDNO – SIGIDOC ID)"];
                    var filename = $"BG_{sealId}";
                    var sequence = sealId.PadLeft(4, '0');
                    allValues.Add("FILENAME", filename);
                    allValues.Add("SEQUENCE", sequence);
                    allValues.Add("{}", "―");

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