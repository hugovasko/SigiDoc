using System.Xml.Linq;
using OfficeOpenXml;

namespace ExcelToXMLConverter
{
    class Program
    {
        static void Main(string[] args)
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                if (worksheet == null)
                {
                    Console.WriteLine("The worksheet does not exist.");
                    return;
                }

                // Get worksheet dimensions
                var dimensions = worksheet.Dimension;

                // Get rows headers from the first column (column A)
                var rowHeaders = worksheet.Cells[1, 1, dimensions.End.Row, 1]
                    .Select(c => c.Value?.ToString().Trim()).ToArray();

                // Find the rows containing the headers we need
                int title_EN_row = Array.IndexOf(rowHeaders, "Title") + 1;
                int editorFN_EN_row = Array.IndexOf(rowHeaders, "Editor forename") + 1;
                int editorSN_EN_row = Array.IndexOf(rowHeaders, "Editor surname") + 1;
                int editionEN_row = Array.IndexOf(rowHeaders, "EDITION(S)") + 1;
                int filename_row = Array.IndexOf(rowHeaders, "Filename") + 1;
                int sealID_row = Array.IndexOf(rowHeaders, "SEAL ID") + 1;
                int sequence_row = Array.IndexOf(rowHeaders, "Sequence") + 1;
                int type_row = Array.IndexOf(rowHeaders, "TYPE") + 1;

                // Loop through all subsequent columns and retrieve the data for each header
                for (int col = 2; col <= dimensions.End.Column; col++)
                {
                    string title_EN = worksheet.Cells[title_EN_row, col].Value?.ToString();
                    string editorFN_EN = worksheet.Cells[editorFN_EN_row, col].Value?.ToString();
                    string editorSN_EN = worksheet.Cells[editorSN_EN_row, col].Value?.ToString();
                    string editionEN = worksheet.Cells[editionEN_row, col].Value?.ToString();
                    string filename = worksheet.Cells[filename_row, col].Value?.ToString();
                    string sealID = worksheet.Cells[sealID_row, col].Value?.ToString();
                    string sequence = worksheet.Cells[sequence_row, col].Value?.ToString();
                    string type = worksheet.Cells[type_row, col].Value?.ToString();

                    // define a dictionary that maps the keys to the corresponding values
                    var allValues = new Dictionary<string, string>
                    {
                        {"{TITLE_EN}", title_EN},
                        {"{FORENAME_EN}", editorFN_EN},
                        {"{SURNAME_EN}", editorSN_EN},
                        {"{EDITION_EN}", editionEN},
                        {"{FILENAME}", filename},
                        {"{SIGIDOC_ID}", sealID},
                        {"{SEQUENCE}", sequence},
                        {"{TYPE}", type}
                    };

                    // replace the XML keys with the corresponding values
                    foreach (var element in doc.Descendants())
                    {
                        if (allValues.TryGetValue(element.Value, out string replacement))
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