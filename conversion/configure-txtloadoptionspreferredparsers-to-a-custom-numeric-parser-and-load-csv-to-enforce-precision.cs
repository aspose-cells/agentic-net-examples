// Title: C# – Load CSV with Aspose.Cells using a custom Decimal parser (TxtLoadOptions.PreferredParsers)
// Description: Demonstrates how to implement an ICustomParser that converts strings to Decimal, assign it to TxtLoadOptions.PreferredParsers for a specific CSV column, enable KeepPrecision and numeric conversion, load the CSV from a memory stream into a Workbook, verify the cell type, and save the result as an XLSX file.
// Keywords: Aspose.Cells custom parser | TxtLoadOptions PreferredParsers | CSV to Excel high precision | Decimal precision C# | KeepPrecision Aspose.Cells | ConvertNumericData CSV | ICustomParser example | load CSV Aspose.Cells .NET | financial data import Aspose
// Common Searches: Aspose.Cells custom numeric parser for CSV | How to preserve decimal precision when loading CSV | TxtLoadOptions PreferredParsers column specific | C# load CSV with high‑precision numbers Aspose | KeepPrecision option Aspose.Cells CSV import
// Developer Intent: Set TxtLoadOptions.PreferredParsers to a Decimal‑based ICustomParser so that selected CSV columns are imported as exact Decimal values without losing precision.
// Use Cases: Import financial CSV files where amounts have many decimal places and must remain exact in the workbook. | Load mixed‑type CSV data (text in column A, high‑precision numbers in column B) by assigning null to the first parser and a Decimal parser to the second. | Prevent truncation or rounding of large numeric strings during CSV‑to‑Excel conversion by enabling KeepPrecision and disabling date conversion.
// AI Prompts: Generate C# code that creates a TxtLoadOptions object for CSV, assigns a custom ICustomParser that parses values to Decimal for column 2, loads the data into a Workbook, and saves it as XLSX while preserving precision. | Explain the role of TxtLoadOptions.PreferredParsers in Aspose.Cells and how it interacts with KeepPrecision, ConvertNumericData, and ConvertDateTimeData settings. | Show how to validate that a cell imported with a custom Decimal parser is of type Numeric and retrieve its Decimal value in C#.

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsCustomParserDemo
{
    // Custom numeric parser that parses values as Decimal to preserve precision
    // Demonstrates how to implement an ICustomParser that converts strings to Decimal, assign it to TxtLoadOptions.PreferredParsers for a specific CSV column, enable KeepPrecision and numeric conversion, load the CSV from a memory stream into a Workbook, verify the cell type, and save the result as an XLSX file.
    class DecimalParser : ICustomParser
    {
        // Attempts to parse the string; returns true if successful
        public bool Parse(string value, out object result)
        {
            // Use invariant culture to avoid locale‑dependent decimal separators
            if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal dec))
            {
                result = dec;
                return true;
            }
            result = null;
            return false;
        }

        // Direct parsing without the out parameter
        public object ParseObject(string value)
        {
            if (decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal dec))
                return dec;
            // Fallback to original string if parsing fails
            return value;
        }

        // Description of the parser format
        public string GetFormat()
        {
            return "Decimal";
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample CSV containing high‑precision numeric values
            string csvData = "ID,Value\n1,1234567890.123456789\n2,9876543210.987654321";

            // Create TxtLoadOptions for CSV and assign the custom parser
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                // PreferredParsers[0] = null (default parser for first column)
                // PreferredParsers[1] = DecimalParser for the second column (Value)
                PreferredParsers = new ICustomParser[] { null, new DecimalParser() },

                // KeepPrecision ensures long string values are not truncated
                KeepPrecision = true,

                // Enable numeric conversion (default true) and disable date conversion for clarity
                ConvertNumericData = true,
                ConvertDateTimeData = false
            };

            // Load the CSV data from a memory stream using the configured options
            using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(csvData)))
            {
                Workbook workbook = new Workbook(stream, loadOptions);
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Verify that the high‑precision number is stored as a numeric type
                Console.WriteLine("Cell B2 Type: " + cells[1, 1].Type); // Expected: Numeric
                Console.WriteLine("Cell B2 Value (Decimal): " + cells[1, 1].Value);

                // Save the workbook to an Excel file
                workbook.Save("CustomParserOutput.xlsx", SaveFormat.Xlsx);
            }
        }
    }
}
