using System;
using System.IO;
using Aspose.Cells;

class PreventExponentialCsvImport
{
    static void Main()
    {
        // Create a sample CSV file with large numeric values if it does not exist
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "large_numbers.csv");
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath, "12345678901234567890,98765432109876543210\r\n");
        }

        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure TxtLoadOptions to keep precision and preserve exact formatting
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            Separator = ',',
            ConvertNumericData = true,
            KeepPrecision = true,
            LoadStyleStrategy = TxtLoadStyleStrategy.ExactFormat
        };

        // Import the CSV data starting at cell A1 (row 0, column 0)
        cells.ImportCSV(csvPath, loadOptions, 0, 0);

        // Save the workbook to an Excel file
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}