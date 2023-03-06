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

                // Load the XML document into memory
                XDocument doc;
                using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                {
                    doc = XDocument.Load(stream);
                }

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
                var titleEnRow = Array.IndexOf(rowHeaders, "Title") + 1;
                var editorFnEnRow = Array.IndexOf(rowHeaders, "Editor forename") + 1;
                var editorSnEnRow = Array.IndexOf(rowHeaders, "Editor surname") + 1;
                var editionEnRow = Array.IndexOf(rowHeaders, "EDITION(S)") + 1;
                var sealIdRow = Array.IndexOf(rowHeaders, "SEAL ID") + 1;
                var typeRow = Array.IndexOf(rowHeaders, "TYPE") + 1;

                // Loop through all subsequent columns and retrieve the data for each header
                for (var col = 2; col <= dimensions.End.Column; col++)
                {
                    // Get the values for each header
                    var titleEn = worksheet.Cells[titleEnRow, col].Value?.ToString() ?? "No data";
                    var editorFnEn = worksheet.Cells[editorFnEnRow, col].Value?.ToString() ?? "No data";
                    var editorSnEn = worksheet.Cells[editorSnEnRow, col].Value?.ToString() ?? "No data";
                    var editionEn = worksheet.Cells[editionEnRow, col].Value?.ToString() ?? "No data";
                    var sealId = worksheet.Cells[sealIdRow, col].Value?.ToString() ?? "No data";
                    var type = worksheet.Cells[typeRow, col].Value?.ToString() ?? "No data";

                    // Generate filename
                    var filename = $"TM_{sealId}";

                    // Generate sequence
                    var sequence = sealId.PadLeft(4, '0');

                    // Define a dictionary that maps the keys to the corresponding values
                    var allValues = new Dictionary<string, string>
                    {
                        {"{TITLE_EN}", titleEn},
                        {"{FORENAME_EN}", editorFnEn},
                        {"{SURNAME_EN}", editorSnEn},
                        {"{EDITION_EN}", editionEn},
                        {"{FILENAME}", filename},
                        {"{SIGIDOC_ID}", sealId},
                        {"{SEQUENCE}", sequence},
                        {"{TYPE}", type}
                    };

                    // Replace the XML keys with the corresponding values
                    foreach (var element in doc.Descendants())
                    {
                        if (allValues.TryGetValue(element.Value, out var replacement))
                        {
                            element.Value = replacement;
                        }
                    }

                    // Save the updated XML file to disk
                    doc.Save($"./resources/{filename}.xml");

                    // Reset the dictionary
                    allValues.Clear();

                    // Reset the XML document
                    doc = XDocument.Load(@"./resources/SigiDocTemplate.xml");
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