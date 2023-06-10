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
                // Set EPPlus license context to NonCommercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Load the XML template file into memory
                XDocument xmlTemplate;
                using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                {
                    xmlTemplate = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                }

                // Load the XML seals list file into memory
                XDocument sealsList;
                using (var stream = new StreamReader(@"../all_seals.xml"))
                {
                    sealsList = XDocument.Load(stream);
                }

                XNamespace ns = "http://www.tei-c.org/ns/1.0";

                // Load the Excel file into memory
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

                // Get worksheet dimensions
                var dimensions = worksheet.Dimension;

                // Define a dictionary to hold the headers and their positions
                var headers = new Dictionary<string, int>();

                // Loop through all rows in the first column and find the ones containing the headers we need
                for (var row = 1; row <= dimensions.End.Row; row++)
                {
                    var header = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                    if (header == null) continue;
                    headers.Add(header, row);
                }

                // Loop through all subsequent columns and retrieve the data for each header
                for (var col = 2; col <= dimensions.End.Column; col++)
                {

                    // Define a dictionary to hold the header and its value
                    var allValues = new Dictionary<string, string>();

                    // Loop through all headers and get their values
                    foreach (var header in headers)
                    {
                        var value = worksheet.Cells[header.Value, col].Value?.ToString()?.Trim() ?? "―";
                        allValues.Add(header.Key, value);
                    }

                    // // Generate a filename and sequence
                    var sealId = allValues["SEAL ID (IDNO – SIGIDOC ID)"];
                    var filename = $"BG_{sealId}";
                    var sequence = sealId.PadLeft(4, '0');

                    // Add the filename and sequence to the dictionary
                    allValues.Add("FILENAME", filename);
                    allValues.Add("SEQUENCE", sequence);

                    // Get not before and not after dates from internal date
                    // var internalDate = allValues["INTERNAL DATE"];
                    // var notBefore = internalDate.Split('-')[0].PadLeft(4, '0');
                    // var notAfter = internalDate.Split('-')[1].PadLeft(4, '0');

                    // Add the not before and not after dates to the dictionary
                    // allValues.Add("NOT BEFORE", notBefore);
                    // allValues.Add("NOT AFTER", notAfter);

                    // Add empty curly braces to the dictionary
                    allValues.Add("{}", "―");



                    // Replace the XML keys with the corresponding values
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

                    // Save the updated XML file to disk
                    xmlTemplate.Save($"../webapps/ROOT/content/xml/epidoc/{filename}.xml");

                    // Reset the dictionary
                    allValues.Clear();

                    // Reset the XML document
                    using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                    {
                        xmlTemplate = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
                    }

                    // Get the <list> element from the seals list xml file
                    var listElement = sealsList.Descendants(ns + "list").FirstOrDefault();

                    if (listElement != null)
                    {
                        // Check if an item with the same "n" attribute value already exists in the seals list
                        bool isItemInList = listElement.Descendants(ns + "item").Any(i => i.Attribute("n")?.Value == filename);

                        if (!isItemInList)
                        {
                            // Create a new list item
                            var newListItem = new XElement(ns + "item");
                            newListItem.SetAttributeValue("n", filename);
                            newListItem.SetAttributeValue("sortKey", sequence);
                            listElement.Add(newListItem);

                            // Sort the list items by their "sortKey" attribute value
                            var sortedListItems = listElement.Descendants(ns + "item").OrderBy(i => i.Attribute("sortKey")?.Value).ToList();

                            // Replace the list items with the sorted ones
                            listElement.ReplaceNodes(sortedListItems);

                            // Save the updated seals list xml file to disk
                            sealsList.Save(@"../all_seals.xml");

                            // Reset the seals list xml document
                            sealsList = XDocument.Load(@"../all_seals.xml");
                        }
                    }
                }

                // Close the Excel file
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