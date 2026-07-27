// Title: C# Custom dd/MM/yyyy Date Parser for CSV Loading with Aspose.Cells PreferredParsers
// Description: Demonstrates how to implement an ICustomParser that recognises dates in the "dd/MM/yyyy" pattern, assign it to TxtLoadOptions.PreferredParsers for the first CSV column, load the data into a Workbook, verify the DateTime values, and save the result as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells custom parser | PreferredParsers CSV | ICustomParser dd/MM/yyyy | C# load CSV Aspose.Cells | European date format parsing | .NET Excel import custom date | TxtLoadOptions custom parser | Aspose.Cells example GitHub
// Common Searches: how to use ICustomParser with Aspose.Cells | parse dd/MM/yyyy dates from CSV in C# | assign custom date parser to PreferredParsers | load CSV with European date format Aspose.Cells | Aspose.Cells custom parser example GitHub
// Developer Intent: Create a custom parser for "dd/MM/yyyy" dates and set it as the PreferredParser for the first column when importing a CSV workbook with Aspose.Cells.
// Use Cases: Import CSV files where dates are stored as dd/MM/yyyy and need to be recognised as DateTime cells. | Combine a custom date parser with default numeric parsers to handle mixed‑type columns during CSV import. | Generate an Excel workbook from CSV data while preserving correct date formatting for downstream reporting.
// AI Prompts: Write a C# ICustomParser that parses "dd/MM/yyyy" strings and integrates it into TxtLoadOptions.PreferredParsers. | Show code to validate that the custom parser correctly populated cell.DateTimeValue after loading a CSV. | Provide a fallback strategy for rows with invalid date strings when using a custom date parser in Aspose.Cells.

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomDateParserDemo
{
    // Custom parser that interprets dates in "dd/MM/yyyy" format
    // Demonstrates how to implement an ICustomParser that recognises dates in the "dd/MM/yyyy" pattern, assign it to TxtLoadOptions.PreferredParsers for the first CSV column, load the data into a Workbook, verify the DateTime values, and save the result as an Excel file using Aspose.Cells for .NET.
    public class DdMMyyyyParser : ICustomParser
    {
        private string _lastFormat;

        // Parses the string and returns a DateTime object if the format matches
        public object ParseObject(string value)
        {
            if (DateTime.TryParseExact(value, "dd/MM/yyyy", CultureInfo.InvariantCulture,
                                       DateTimeStyles.None, out DateTime dt))
            {
                _lastFormat = "dd/MM/yyyy";
                return dt;
            }

            _lastFormat = null;
            return null;
        }

        // Required by ICustomParser – delegates to ParseObject
        public bool Parse(string value, out object result)
        {
            result = ParseObject(value);
            return result != null;
        }

        // Returns the format used for the last successful parse
        public string GetFormat()
        {
            return _lastFormat;
        }
    }

    class Program
    {
        static void Main()
        {
            // Sample CSV data where the first column contains dates in dd/MM/yyyy format
            string csvData = "01/02/2023,123.45\n15/03/2023,678.90";

            // Convert the CSV string to a memory stream
            byte[] csvBytes = System.Text.Encoding.UTF8.GetBytes(csvData);
            using var csvStream = new MemoryStream(csvBytes);

            // Create TxtLoadOptions for CSV format
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);

            // Assign the custom date parser to the first column (index 0)
            loadOptions.PreferredParsers = new ICustomParser[] { new DdMMyyyyParser(), null };

            // Load the workbook using the custom parser
            Workbook workbook = new Workbook(csvStream, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];

            // Demonstrate that the first column is parsed as DateTime
            Console.WriteLine("A1 (Date)   : " + sheet.Cells[0, 0].DateTimeValue.ToString("yyyy-MM-dd"));
            Console.WriteLine("A2 (Date)   : " + sheet.Cells[1, 0].DateTimeValue.ToString("yyyy-MM-dd"));

            // Second column remains numeric (default parser)
            Console.WriteLine("B1 (Number) : " + sheet.Cells[0, 1].DoubleValue);
            Console.WriteLine("B2 (Number) : " + sheet.Cells[1, 1].DoubleValue);

            // Save the result to an Excel file
            workbook.Save("CustomDateParserOutput.xlsx");
        }
    }
}
