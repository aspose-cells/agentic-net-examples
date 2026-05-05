using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvLoadExample
{
    // Custom parser that forces the first column to be treated as a string
    public class StringParser : ICustomParser
    {
        public bool Parse(string value, out object result)
        {
            // Always return the raw string value
            result = value;
            return true;
        }

        public object ParseObject(string value)
        {
            // Return the raw string value
            return value;
        }

        public string GetFormat()
        {
            // Description of the parser format
            return "String";
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to a temporary CSV file
            string csvPath = Path.Combine(Path.GetTempPath(), "sample.csv");

            // Create sample CSV content
            File.WriteAllText(csvPath,
                "\"2023-01-01\",100\n\"2023-01-02\",200\n\"2023-01-03\",300");

            // ------------------------------------------------------------
            // 1. Create TxtLoadOptions for CSV loading
            // ------------------------------------------------------------
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.Separator = ',';                 // Use comma as delimiter
            loadOptions.HasTextQualifier = true;         // Respect quoted values

            // 2. Assign a custom parser for the first column (index 0)
            //    The parser will treat the first column as a plain string.
            loadOptions.PreferredParsers = new ICustomParser[]
            {
                new StringParser(), // Column 0 parser
                null                // Remaining columns use default parsers
            };

            // ------------------------------------------------------------
            // 3. Load the CSV file into a Workbook using the configured options
            // ------------------------------------------------------------
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // ------------------------------------------------------------
            // 4. Demonstrate that the first column is stored as a string
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("A1 (string)   : " + sheet.Cells["A1"].StringValue);
            Console.WriteLine("B1 (numeric)  : " + sheet.Cells["B1"].IntValue);
            Console.WriteLine("A2 (string)   : " + sheet.Cells["A2"].StringValue);
            Console.WriteLine("B2 (numeric)  : " + sheet.Cells["B2"].IntValue);

            // ------------------------------------------------------------
            // 5. Save the workbook as an XLSX file
            // ------------------------------------------------------------
            string outputPath = Path.Combine(Path.GetTempPath(), "CsvLoaded.xlsx");
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to: {outputPath}");

            // Cleanup temporary CSV file
            if (File.Exists(csvPath))
                File.Delete(csvPath);
        }
    }
}