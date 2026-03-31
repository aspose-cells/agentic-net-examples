using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the temporary CSV file
            string csvPath = "large_numbers.csv";

            // Create sample CSV content with large numeric values
            // Values are longer than 15 digits to test precision handling
            string csvContent = "ID,Value\n" +
                                "1,123456789012345\n" +
                                "2,98765432109876543210\n" +
                                "3,0.0000001234567890123456";

            // Write the CSV content to the file using UTF-8 encoding
            File.WriteAllText(csvPath, csvContent, Encoding.UTF8);

            // Configure text load options for CSV
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // Use comma as the separator
                Separator = ',',
                // Do NOT convert numeric strings to numeric data; keep them as text
                ConvertNumericData = false,
                // Preserve precision for long strings (prevents scientific notation)
                KeepPrecision = true,
                // Optional: keep original formatting style (not required for this task)
                LoadStyleStrategy = TxtLoadStyleStrategy.None
            };

            // Load the CSV file into a workbook using the specified options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Access the first worksheet to demonstrate that values are retained as strings
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A2 (ID): " + sheet.Cells["A2"].StringValue);
            Console.WriteLine("Cell B2 (Large Number): " + sheet.Cells["B2"].StringValue);
            Console.WriteLine("Cell B3 (Very Small Number): " + sheet.Cells["B3"].StringValue);

            // Save the workbook to an Excel file
            string outputPath = "LargeNumbersOutput.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");

            // Clean up the temporary CSV file
            if (File.Exists(csvPath))
                File.Delete(csvPath);
        }
    }
}