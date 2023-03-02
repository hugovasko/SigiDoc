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
                // Load the XML document into memory
                XDocument doc;
                using (var stream = new StreamReader(@"./resources/SigiDocTemplate.xml"))
                {
                    doc = XDocument.Load(stream);
                }

                // Set EPPlus license context to NonCommercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

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

                string title_EN = worksheet.Cells["B57"].Value?.ToString();
                string editorFN_EN = worksheet.Cells["B58"].Value?.ToString();
                string editorSN_EN = worksheet.Cells["B59"].Value?.ToString();
                string edition = worksheet.Cells["B45"].Value?.ToString();
                string filename = worksheet.Cells["B60"].Value?.ToString();
                string sequence = worksheet.Cells["B61"].Value?.ToString();
                string id = worksheet.Cells["B1"].Value?.ToString();
                string type = worksheet.Cells["B2"].Value?.ToString();

                // define the dictionary that maps the keys to the corresponding editor values
                var allValues = new Dictionary<string, string>
                {
                    {"{TITLE_EN}", title_EN},
                    {"{FORENAME_EN}", editorFN_EN},
                    {"{SURNAME_EN}", editorSN_EN},
                    {"{EDITION}", edition},
                    {"{FILENAME}", filename},
                    {"{SIGIDOC_ID}", id},
                    {"{SEQUENCE}", sequence},
                    {"{TYPE}", type}
                };

                // replace the keys with the corresponding values
                foreach (var element in doc.Descendants())
                {
                    if (allValues.TryGetValue(element.Value, out string replacement))
                    {
                        element.Value = replacement;
                    }
                }

                // Save the updated XML file to disk
                doc.Save($"./resources/{filename}.xml");

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