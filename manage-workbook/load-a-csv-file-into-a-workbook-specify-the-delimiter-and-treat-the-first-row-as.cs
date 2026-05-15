using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvImport
{
    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be imported
            string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "input.csv");

            // Ensure the CSV file exists (create a sample if it does not)
            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath, "Header1;Header2\n123;456\nABC;DEF");
            }

            // Create TxtLoadOptions for CSV import
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = ';',
                HeaderRowsCount = 1,
                ConvertNumericData = true
            };

            // Load the CSV file into a workbook using the specified options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Access the first worksheet to verify data
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"First cell value: {sheet.Cells["A1"].StringValue}");
            Console.WriteLine($"Second cell value: {sheet.Cells["B2"].StringValue}");

            // Save the workbook to an Excel file
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}