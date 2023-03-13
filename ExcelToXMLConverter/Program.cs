using System.Xml.Linq;
using OfficeOpenXml;

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
                    xmlTemplate = XDocument.Load(stream);
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
                using (var stream = new FileStream(@"./resources/SIGIDOC CELLS ENG.xlsx", FileMode.Open, FileAccess.Read))
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

                // Get rows headers from the first column (column A)
                var rowHeaders = worksheet.Cells[1, 1, dimensions.End.Row, 1]
                    .Select(c => c.Value?.ToString()?.Trim()).ToArray();

                // Find the rows containing the headers we need
                var titleEnRow = Array.IndexOf(rowHeaders, "TITLE") + 1;
                var editorFnEnRow = Array.IndexOf(rowHeaders, "EDITOR FORENAME") + 1;
                var editorSnEnRow = Array.IndexOf(rowHeaders, "EDITOR SURNAME") + 1;
                var editionEnRow = Array.IndexOf(rowHeaders, "EDITION(S)") + 1;
                var sealIdRow = Array.IndexOf(rowHeaders, "SEAL ID") + 1;
                var typeEnRow = Array.IndexOf(rowHeaders, "TYPE") + 1;
                var findPlaceEnRow = Array.IndexOf(rowHeaders, "FIND PLACE") + 1;
                var dateRow = Array.IndexOf(rowHeaders, "DATE") + 1;
                var internalDateRow = Array.IndexOf(rowHeaders, "INTERNAL DATE") + 1;
                var generalLayoutEnRow = Array.IndexOf(rowHeaders, "GENERAL LAYOUT") + 1;
                var typeOfImpressionEnRow = Array.IndexOf(rowHeaders, "TYPE OF IMPRESSION") + 1;
                var materialEnRow = Array.IndexOf(rowHeaders, "MATERIAL") + 1;
                var shapeEnRow = Array.IndexOf(rowHeaders, "SHAPE") + 1;
                var diameterRow = Array.IndexOf(rowHeaders, "DIMENSIONS (mm)") + 1;
                var datingCriteriaEnRow = Array.IndexOf(rowHeaders, "DATING CRITERIA") + 1;
                var alternativeDatingRow = Array.IndexOf(rowHeaders, "ALTERNATIVE DATING") + 1;

                // Loop through all subsequent columns and retrieve the data for each header
                for (var col = 2; col <= dimensions.End.Column; col++)
                {
                    // Get the values for each header
                    var titleEn = worksheet.Cells[titleEnRow, col].Value?.ToString() ?? "-";
                    var editorFnEn = worksheet.Cells[editorFnEnRow, col].Value?.ToString() ?? "-";
                    var editorSnEn = worksheet.Cells[editorSnEnRow, col].Value?.ToString() ?? "-";
                    var editionEn = worksheet.Cells[editionEnRow, col].Value?.ToString() ?? "-";
                    var sealId = worksheet.Cells[sealIdRow, col].Value?.ToString() ?? "-";
                    var typeEn = worksheet.Cells[typeEnRow, col].Value?.ToString() ?? "-";
                    var findPlaceEn = worksheet.Cells[findPlaceEnRow, col].Value?.ToString() ?? "-";
                    var date = worksheet.Cells[dateRow, col].Value?.ToString() ?? "-";
                    var internalDate = worksheet.Cells[internalDateRow, col].Value?.ToString() ?? "-";
                    var generalLayoutEn = worksheet.Cells[generalLayoutEnRow, col].Value?.ToString() ?? "-";
                    var typeOfImpressionEn = worksheet.Cells[typeOfImpressionEnRow, col].Value?.ToString() ?? "-";
                    var materialEn = worksheet.Cells[materialEnRow, col].Value?.ToString() ?? "-";
                    var shapeEn = worksheet.Cells[shapeEnRow, col].Value?.ToString() ?? "-";
                    var diameter = worksheet.Cells[diameterRow, col].Value?.ToString() ?? "-";
                    var datingCriteriaEn = worksheet.Cells[datingCriteriaEnRow, col].Value?.ToString() ?? "-";
                    var alternativeDating = worksheet.Cells[alternativeDatingRow, col].Value?.ToString() ?? "-";

                    // Generate filename
                    var filename = $"TM_{sealId}";

                    // Generate sequence
                    var sequence = sealId.PadLeft(4, '0');

                    // Get not before and not after dates from internal date
                    var notBefore = internalDate.Split('-')[0].PadLeft(4, '0');
                    var notAfter = internalDate.Split('-')[1].PadLeft(4, '0');

                    // Define a dictionary that maps the keys to the corresponding values
                    var allValues = new Dictionary<string, string>
                    {
                        {"{TITLE_EN}", titleEn},
                        {"{EDITOR_FORENAME_EN}", editorFnEn},
                        {"{EDITOR_SURNAME_EN}", editorSnEn},
                        {"{EDITION_EN}", editionEn},
                        {"{FILENAME}", filename},
                        {"{SIGIDOC_ID}", sealId},
                        {"{SEQUENCE}", sequence},
                        {"{TYPE_EN}", typeEn},
                        {"{FIND_PLACE_EN}", findPlaceEn},
                        {"{DATE}", date},
                        {"{INTERNAL_DATE}", internalDate},
                        {"{GENERAL_LAYOUT_EN}", generalLayoutEn},
                        {"{TYPE_OF_IMPRESSION_EN}", typeOfImpressionEn},
                        {"{MATERIAL_EN}", materialEn},
                        {"{SHAPE_EN}", shapeEn},
                        {"{DIAMETER}", diameter},
                        {"{DATING_CRITERIA_EN}", datingCriteriaEn},
                        {"{ALTERNATIVE_DATING}", alternativeDating},
                        {"{NOT_BEFORE}", notBefore},
                        {"{NOT_AFTER}", notAfter},
                        {"{}", "-"}
                    };

                    // Replace the XML keys with the corresponding values
                    foreach (var element in xmlTemplate.Descendants())
                    {
                        if (allValues.TryGetValue(element.Value, out var elementReplacement))
                        {
                            element.Value = elementReplacement;
                        }
                        foreach (var attribute in element.Attributes())
                        {
                            if (allValues.TryGetValue(attribute.Value, out var attributeReplacement))
                            {
                                attribute.Value = attributeReplacement;
                            }
                        }
                    }

                    // Save the updated XML file to disk
                    xmlTemplate.Save($"./resources/{filename}.xml");


                    // Reset the dictionary
                    allValues.Clear();

                    // Reset the XML document
                    xmlTemplate = XDocument.Load(@"./resources/SigiDocTemplate.xml");

                    // Get the <list> element from the seals list xml file
                    var listElement = sealsList.Descendants(ns + "list").FirstOrDefault();

                    if (listElement != null)
                    {
                        // Create a new list item
                        var newListItem = new XElement(ns + "item");
                        newListItem.SetAttributeValue("n", filename);
                        newListItem.SetAttributeValue("sortKey", sequence);
                        listElement.Add(newListItem);

                        // Save the updated seals list xml file to disk
                        sealsList.Save(@"../all_seals.xml");

                        // Reset the seals list xml document
                        sealsList = XDocument.Load(@"../all_seals.xml");
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