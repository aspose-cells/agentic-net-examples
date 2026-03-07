using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToExcelConversion
{
    class Program
    {
        static void Main()
        {
            // Paths for the source CSV file and the destination Excel file
            string csvPath = "sample.csv";
            string excelPath = "converted.xlsx";

            // Create a sample CSV file (replace with your actual CSV/JSON content if needed)
            string csvContent = "Name,Age,City\nJohn,30,New York\nAlice,25,London\nBob,35,Paris";
            File.WriteAllText(csvPath, csvContent);

            try
            {
                // LoadOptions specifying that the source file is a CSV
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // SaveOptions for saving as XLSX (OOXML format)
                SaveOptions saveOptions = new OoxmlSaveOptions();

                // Convert CSV to Excel using the Aspose.Cells ConversionUtility
                ConversionUtility.Convert(csvPath, loadOptions, excelPath, saveOptions);

                Console.WriteLine($"Conversion successful: '{csvPath}' -> '{excelPath}'");

                // Optional verification: load the resulting workbook and display sheet count
                Workbook wb = new Workbook(excelPath);
                Console.WriteLine($"Workbook contains {wb.Worksheets.Count} worksheet(s).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
            finally
            {
                // Clean up the temporary CSV file (keep the Excel file if needed)
                if (File.Exists(csvPath))
                {
                    File.Delete(csvPath);
                }
            }
        }
    }
}