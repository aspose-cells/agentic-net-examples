// Title: C# – Load CSV with Aspose.Cells and a user‑defined decimal parser to retain exact precision
// Description: Demonstrates implementing ICustomParser that returns System.Decimal, assigning it via TxtLoadOptions.PreferredParsers for a target CSV column, enabling KeepPrecision and ConvertNumericData, loading the stream into a Workbook, confirming the numeric type, and saving the result as an XLSX file.
// Keywords: Aspose.Cells | TxtLoadOptions | PreferredParsers | custom numeric parser | decimal precision | CSV import .NET | KeepPrecision | ConvertNumericData | C# example | financial data import
// Common Searches: Aspose.Cells custom parser CSV | How to preserve decimal places when loading CSV with Aspose.Cells | Set PreferredParsers for specific columns in C# | Enable KeepPrecision for large numbers Aspose.Cells | Load high‑precision numbers into Excel using Aspose.Cells .NET
// Developer Intent: Configure TxtLoadOptions so that a selected CSV column is parsed with a user‑defined decimal parser, guaranteeing high‑precision values are stored as Decimal in the workbook.
// Use Cases: Import financial statements where amounts exceed 15 decimal places. | Read scientific measurement data that requires exact numeric values. | Apply distinct parsers per column while retaining default handling for the rest. | Generate Excel reports from CSV without losing numeric accuracy.
// AI Prompts: Generate C# code that uses ICustomParser to parse dates in a CSV with Aspose.Cells. | Show how to assign separate parsers for three columns using TxtLoadOptions.PreferredParsers. | Explain the interaction between KeepPrecision and ConvertNumericData for large numeric strings.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCustomParserDemo
{
    // Custom numeric parser that preserves full decimal precision
    // Demonstrates implementing ICustomParser that returns System.Decimal, assigning it via TxtLoadOptions.PreferredParsers for a target CSV column, enabling KeepPrecision and ConvertNumericData, loading the stream into a Workbook, confirming the numeric type, and saving the result as an XLSX file.
    public class DecimalParser : ICustomParser
    {
        // Try to parse the string as a decimal. Return true if successful.
        public bool Parse(string value, out object result)
        {
            // Trim whitespace and try to parse using invariant culture
            if (decimal.TryParse(value.Trim(), System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, out decimal dec))
            {
                result = dec; // Return the decimal value
                return true;
            }

            // Fallback to default parsing (treat as string)
            result = value;
            return false;
        }

        // Direct parsing without the out parameter (used internally by Aspose.Cells)
        public object ParseObject(string value)
        {
            if (decimal.TryParse(value.Trim(), System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, out decimal dec))
            {
                return dec;
            }
            return value;
        }

        // Description of the format this parser handles
        public string GetFormat()
        {
            return "Decimal";
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample CSV data containing high‑precision numbers
            string csvData = "ID,Amount,Description\n" +
                             "1,12345.67890123456789,First item\n" +
                             "2,98765.43210987654321,Second item";

            // Create TxtLoadOptions for CSV and assign the custom decimal parser
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // Use the custom parser for the second column (index 1)
                PreferredParsers = new ICustomParser[] { null, new DecimalParser() },

                // Keep precision for long numeric strings
                KeepPrecision = true,

                // Ensure numeric conversion is enabled
                ConvertNumericData = true,

                // Optional: keep date conversion disabled for this example
                ConvertDateTimeData = false
            };

            // Load the CSV data into a workbook using the configured options
            using (MemoryStream csvStream = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(csvStream, loadOptions);
                Cells cells = workbook.Worksheets[0].Cells;

                // Demonstrate that the high‑precision numbers are stored as Decimal
                Console.WriteLine("Cell B2 Type: " + cells[1, 1].Type); // Should be Numeric
                Console.WriteLine("Cell B2 Value: " + cells[1, 1].Value); // Full precision

                // Save the workbook to an Excel file
                workbook.Save("PreciseNumbers.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}
