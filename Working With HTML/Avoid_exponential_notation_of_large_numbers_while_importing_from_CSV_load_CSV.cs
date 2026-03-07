using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Resolve CSV file path relative to the executable directory
            string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "large_numbers.csv");

            // If the CSV does not exist, create a sample file
            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath, "12345678901234567890,0.000000123456789\n98765432109876543210,12345678901234567890");
            }

            // Configure TxtLoadOptions to preserve the exact format of numeric values
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                ConvertNumericData = true,
                LoadStyleStrategy = TxtLoadStyleStrategy.ExactFormat,
                KeepPrecision = true
            };

            // Import the CSV data starting at cell A1 (row 0, column 0)
            cells.ImportCSV(csvPath, loadOptions, 0, 0);

            // Save the workbook to an Excel file
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "output.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("CSV imported successfully without exponential notation.");
        }
    }
}