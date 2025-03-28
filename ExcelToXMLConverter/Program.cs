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
                using (var stream = new FileStream(@"./resources/SIGIDOC_ENG_BG_2025.03.26.test.xlsx", FileMode.Open, FileAccess.Read))
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
                        if ((header.Key == "ANALYSIS DATE NOT BEFORE" || header.Key == "ANALYSIS DATE NOT AFTER") && value.Length < 4)
                        {
                            value = value.PadLeft(4, '0');
                        }
                        allValues.Add(header.Key, value);
                    }

                    var sealId = allValues["SIGIDOC ID"];
                    var filename = allValues["FILENAME"];
                    var sequence = sealId.PadLeft(4, '0');
                    allValues.Add("SEQUENCE", sequence);
                    allValues.Add("{}", "―");

                    XmlUtils.ProcessInterpretiveOrDiplomaticText(xmlTemplate, ns, allValues, "EDITION INTERPRETIVE", "edition", "editorial", "obv");
                    XmlUtils.ProcessInterpretiveOrDiplomaticText(xmlTemplate, ns, allValues, "EDITION DIPLOMATIC", "edition", "diplomatic", "obv");

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

                    var prettyXml = XmlUtils.Prettify(xmlTemplate.ToString());
                    File.WriteAllText($"../webapps/ROOT/content/xml/epidoc/{filename}.xml", prettyXml);
                    allValues.Clear();

                    using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                    {
                        xmlTemplate = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                    }

                    XmlUtils.UpdateSealsList(ns, filename, sequence, sealsList);
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