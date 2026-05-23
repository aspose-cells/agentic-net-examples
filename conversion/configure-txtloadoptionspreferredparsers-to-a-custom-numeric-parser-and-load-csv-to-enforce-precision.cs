using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomParserDemo
{
    // Custom numeric parser that parses values as high‑precision decimal numbers
    class HighPrecisionNumericParser : ICustomParser
    {
        // Attempts to parse the string; returns true if successful and sets the result
        public bool Parse(string value, out object result)
        {
            if (decimal.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out decimal dec))
            {
                result = dec;               // Store as decimal to keep precision
                return true;
            }
            result = null;
            return false;
        }

        // Direct parsing without a success flag; falls back to original string if parsing fails
        public object ParseObject(string value)
        {
            if (decimal.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out decimal dec))
            {
                return dec;
            }
            return value;
        }

        // Description of the parser's format
        public string GetFormat()
        {
            return "HighPrecisionDecimal";
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample CSV data: first column contains high‑precision numbers
            string csvData = "123456789012345.6789012345,Alpha\n" +
                             "987654321098765.4321098765,Beta";

            // Create TxtLoadOptions for CSV loading
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // Use the custom parser for the first column; null for others (default parser)
                PreferredParsers = new ICustomParser[] { new HighPrecisionNumericParser(), null },
                // Keep precision for long strings (optional but often useful)
                KeepPrecision = true,
                // Ensure numeric conversion is enabled
                ConvertNumericData = true,
                // Convert date strings if any (kept true by default)
                ConvertDateTimeData = true
            };

            // Load the CSV from a memory stream using the configured options
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(stream, loadOptions);

                // Demonstrate that the first column retains high precision
                Console.WriteLine("A1 Type : " + workbook.Worksheets[0].Cells[0, 0].Type);
                Console.WriteLine("A1 Value: " + workbook.Worksheets[0].Cells[0, 0].Value);
                Console.WriteLine("B1 Value: " + workbook.Worksheets[0].Cells[0, 1].StringValue);

                // Save the workbook to an Excel file
                workbook.Save("CustomParserOutput.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}