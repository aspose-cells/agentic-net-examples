using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCsvParsingDemo
{
    // Custom parser that attempts to parse a string to double.
    // If parsing fails, it returns false to indicate failure.
    // This parser does not throw, but we will simulate an exception
    // in the loading process by using an invalid parser implementation.
    public class SafeNumberParser : ICustomParser
    {
        public bool Parse(string value, out object result)
        {
            // Try to parse numeric value; if it fails, treat as string.
            if (double.TryParse(value, out double num))
            {
                result = num;
                return true;
            }

            result = value;
            return true; // Always succeed, returning original string.
        }

        public object ParseObject(string value)
        {
            // This method is used by Aspose.Cells during loading.
            // Throw an exception for a specific sentinel value to demonstrate error handling.
            if (value == "##ERROR##")
                throw new InvalidOperationException("Custom parsing error for sentinel value.");

            if (double.TryParse(value, out double num))
                return num;

            return value;
        }

        public string GetFormat()
        {
            return "General";
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample CSV data containing a value that will trigger a parsing exception.
            string csvData = "Name,Score\nAlice,85\nBob,##ERROR##\nCharlie,92";

            // Prepare a memory stream from the CSV string.
            using (MemoryStream csvStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvData)))
            {
                // Configure TxtLoadOptions for CSV loading.
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    // Use the custom parser for the second column (Score).
                    PreferredParsers = new ICustomParser[] { null, new SafeNumberParser() }
                };

                Workbook workbook = null;

                try
                {
                    // Attempt to load the CSV with the custom parsers.
                    workbook = new Workbook(csvStream, loadOptions);
                    Console.WriteLine("CSV loaded successfully.");
                }
                catch (Exception ex)
                {
                    // Catch any parsing exceptions and log details.
                    Console.WriteLine($"Error while parsing CSV: {ex.Message}");
                }

                if (workbook != null)
                {
                    // Access the first worksheet to demonstrate that data was loaded.
                    Worksheet sheet = workbook.Worksheets[0];
                    Cells cells = sheet.Cells;

                    // Output the loaded values.
                    Console.WriteLine($"A2 (Name): {cells["A2"].StringValue}");
                    Console.WriteLine($"B2 (Score): {cells["B2"].StringValue}");
                    Console.WriteLine($"A3 (Name): {cells["A3"].StringValue}");
                    Console.WriteLine($"B3 (Score): {cells["B3"].StringValue}");

                    // Save the workbook to an Excel file.
                    workbook.Save("ParsedCsvOutput.xlsx", SaveFormat.Xlsx);
                    Console.WriteLine("Workbook saved as ParsedCsvOutput.xlsx");
                }
            }
        }
    }
}