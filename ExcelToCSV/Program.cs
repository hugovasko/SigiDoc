using System.Text;
using ExcelDataReader;

namespace ExcelToCSV;
class Program
{
    static void Main(string[] args)
    {
        // #region get file path
        Console.WriteLine("Enter the path of the excel file: ");
        string excelFilePath = Console.ReadLine();
        Console.WriteLine("Enter the path where the csv file to be saved: ");
        string destinationCsvFilePath = Console.ReadLine();

        // #region convert excel to csv
        Console.WriteLine("Exporting file to CSV...." + "\n");
        ExcelFileHelper.SaveAsCsv(excelFilePath, destinationCsvFilePath);
        Console.WriteLine("File exported to CSV!" + "\n");
        // #endregion
    }
}

public class ExcelFileHelper
{
    public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath)
    {

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            IExcelDataReader reader = null;
            if (excelFilePath.EndsWith(".xls"))
            {
                reader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else if (excelFilePath.EndsWith(".xlsx"))
            {
                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            if (reader == null)
                return false;

            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = false
                }
            });

            var csvContent = string.Empty;
            int row_no = 0;
            while (row_no < ds.Tables[0].Rows.Count)
            {
                var arr = new List<string>();
                for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                {
                    arr.Add(ds.Tables[0].Rows[row_no][i].ToString());
                }
                row_no++;
                csvContent += string.Join(",", arr) + "\n";
            }
            StreamWriter csv = new StreamWriter(destinationCsvFilePath, false);
            csv.Write(csvContent);
            csv.Close();
            return true;
        }
    }
}
