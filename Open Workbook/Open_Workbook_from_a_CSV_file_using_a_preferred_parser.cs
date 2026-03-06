using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvPreferredParserDemo
{
    // Custom parser that forces every value to be treated as a string.
    class StringParser : ICustomParser
    {
        public bool Parse(string value, out object result)
        {
            result = value; // Return the original text.
            return true;    // Indicate parsing succeeded.
        }

        public object ParseObject(string value)
        {
            return value;   // Same as above, for object‑based parsing.
        }

        public string GetFormat()
        {
            return "String"; // Description of the parser.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the CSV file to be loaded.
            string csvPath = "sample.csv";

            // Create a simple CSV file for demonstration purposes.
            File.WriteAllText(csvPath, "\"2021-01-01\",100\n\"2021-01-02\",200");

            // Create TxtLoadOptions for CSV format.
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
            loadOptions.Separator = ','; // Explicitly set the delimiter.

            // Assign preferred parsers:
            //   - First column uses the custom StringParser (treated as text).
            //   - Subsequent columns use the default parsers (null entry).
            loadOptions.PreferredParsers = new ICustomParser[] { new StringParser(), null };

            // Load the workbook using the CSV file and the configured options.
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Access the first worksheet to verify the data.
            Worksheet ws = workbook.Worksheets[0];
            Console.WriteLine($"A1: {ws.Cells["A1"].StringValue} (type: {ws.Cells["A1"].Type})");
            Console.WriteLine($"B1: {ws.Cells["B1"].IntValue} (type: {ws.Cells["B1"].Type})");

            // Save the workbook as an XLSX file.
            workbook.Save("Converted.xlsx", SaveFormat.Xlsx);
        }
    }
}