using System;
using System.IO;
using Aspose.Cells;

class AvoidScientificNotation
{
    static void Main()
    {
        // Path to the CSV file containing large numbers
        string csvPath = "large_numbers.csv";

        // Create sample CSV if it does not exist
        if (!File.Exists(csvPath))
        {
            File.WriteAllText(csvPath, "12345678901234567890\r\n");
        }

        // Configure load options:
        // - Do not convert numeric strings to numbers (keep them as text)
        // - Keep precision for long strings (prevents parsing numbers longer than 15 digits)
        // - Preserve the exact format of the original text
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
        {
            ConvertNumericData = false,
            KeepPrecision = true,
            LoadStyleStrategy = TxtLoadStyleStrategy.ExactFormat
        };

        // Create a new workbook and import the CSV using the configured options
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells.ImportCSV(csvPath, loadOptions, 0, 0);

        // Verify that a large number is stored as text, not in exponential notation
        Console.WriteLine("A1 raw value: " + worksheet.Cells["A1"].StringValue);
        Console.WriteLine("A1 cell type: " + worksheet.Cells["A1"].Type);

        // Save the workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}