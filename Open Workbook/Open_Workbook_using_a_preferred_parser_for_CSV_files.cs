using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvPreferredParserDemo
{
    // Custom parser that forces the cell value to be treated as a string
    class StringParser : ICustomParser
    {
        // Try to parse the value; always succeed and return the original string
        public bool Parse(string value, out object result)
        {
            result = value;
            return true;
        }

        // Direct parsing without success flag (used internally by Aspose.Cells)
        public object ParseObject(string value)
        {
            return value;
        }

        // Description of the parser format
        public string GetFormat()
        {
            return "String";
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be loaded
            string csvPath = "sample.csv";

            // Ensure the sample CSV exists (for demonstration purposes)
            if (!File.Exists(csvPath))
            {
                File.WriteAllText(csvPath, "Name,Age\nJohn,30\nJane,25");
            }

            // Create TxtLoadOptions for CSV loading
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.Separator = ',';                 // Use comma as delimiter
            loadOptions.ConvertNumericData = true;       // Convert numeric strings to numbers where applicable
            loadOptions.ConvertDateTimeData = true;      // Convert date strings to DateTime where applicable

            // Set preferred parsers:
            // - First column (Name) will use the custom StringParser, forcing it to stay as string.
            // - Second column (Age) will use the default parser (null entry).
            loadOptions.PreferredParsers = new ICustomParser[] { new StringParser(), null };

            // Load the workbook using the preferred parser options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Access the first worksheet and display loaded values
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"A1 (Name): {sheet.Cells["A1"].StringValue}"); // Expected "Name"
            Console.WriteLine($"B1 (Age): {sheet.Cells["B1"].StringValue}"); // Expected "Age"
            Console.WriteLine($"A2 (First Name): {sheet.Cells["A2"].StringValue}"); // Expected "John"
            Console.WriteLine($"B2 (First Age) Type: {sheet.Cells["B2"].Type}");   // Expected Numeric

            // Optionally save the workbook as XLSX to verify conversion
            workbook.Save("ConvertedFromCsv.xlsx", SaveFormat.Xlsx);
        }
    }
}