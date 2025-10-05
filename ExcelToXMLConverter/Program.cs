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
                using (var stream = new FileStream(@"./resources/SIGIDOC_ENG_BG_04.10.2025.xlsx", FileMode.Open, FileAccess.Read))
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

                // Read all headers from column A
                for (var row = 1; row <= dimensions.End.Row; row++)
                {
                    var header = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                    if (header == null) continue;
                    headers.Add(header, row);
                }

                // Process each column (each column = one seal)
                for (var col = 2; col <= dimensions.End.Column; col++)
                {
                    var allValues = new Dictionary<string, string>();

                    // Extract all values for this seal
                    foreach (var header in headers)
                    {
                        var value = worksheet.Cells[header.Value, col].Value?.ToString()?.Trim() ?? "―";
                        
                        // Pad dates to 4 digits (e.g., "945" becomes "0945")
                        if ((header.Key == "ANALYSIS DATE NOT BEFORE" || header.Key == "ANALYSIS DATE NOT AFTER") && value.Length < 4 && value != "―")
                        {
                            value = value.PadLeft(4, '0');
                        }
                        
                        allValues.Add(header.Key, value);
                    }

                    var sealId = allValues["SIGIDOC ID"];
                    var filename = allValues["FILENAME"];
                    var sequence = sealId.PadLeft(4, '0');
                    allValues.Add("SEQUENCE", sequence);
                    
                    // Add placeholder columns that might not exist in Excel
                    if (!allValues.ContainsKey("HEIGHT")) allValues.Add("HEIGHT", "―");
                    if (!allValues.ContainsKey("WIDTH")) allValues.Add("WIDTH", "―");
                    if (!allValues.ContainsKey("DEPTH")) allValues.Add("DEPTH", "―");
                    
                    // Fallback for empty placeholder
                    allValues.Add("{}", "―");

                    // Process editorial and diplomatic editions
                    XmlUtils.ProcessInterpretiveOrDiplomaticText(xmlTemplate, ns, allValues, "EDITION INTERPRETIVE", "edition", "editorial", "obv");
                    XmlUtils.ProcessInterpretiveOrDiplomaticText(xmlTemplate, ns, allValues, "EDITION DIPLOMATIC", "edition", "diplomatic", "obv");

                    // Replace all placeholders in template
                    foreach (var element in xmlTemplate.Descendants())
                    {
                        foreach (var replacement in ReplacementValues.Replacements)
                        {
                            // Replace in element values
                            if (element.Value == replacement.key)
                            {
                                element.Value = allValues.ContainsKey(replacement.value) 
                                    ? allValues[replacement.value] 
                                    : "―";
                            }

                            // Replace in attributes
                            foreach (var attribute in element.Attributes())
                            {
                                if (attribute.Value == replacement.key)
                                {
                                    attribute.Value = allValues.ContainsKey(replacement.value) 
                                        ? allValues[replacement.value] 
                                        : "―";
                                }
                            }
                        }
                    }

                    // Save the generated XML file
                    var prettyXml = XmlUtils.Prettify(xmlTemplate.ToString());
                    File.WriteAllText($"../webapps/ROOT/content/xml/epidoc/{filename}.xml", prettyXml);
                    
                    Console.WriteLine($"Generated: {filename}.xml (SigiDoc ID: {sealId})");
                    
                    allValues.Clear();

                    // Reload template for next seal
                    using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                    {
                        xmlTemplate = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                    }

                    // Update the seals list
                    XmlUtils.UpdateSealsList(ns, filename, sequence, sealsList);
                }
                
                package.Dispose();
                Console.WriteLine("\n=== SUCCESS! All seals generated. ===");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
    }
}