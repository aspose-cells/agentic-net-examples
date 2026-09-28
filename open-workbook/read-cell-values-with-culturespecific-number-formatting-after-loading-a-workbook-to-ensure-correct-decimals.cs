// Title: Read a numeric cell with German locale formatting using Aspose.Cells for .NET
// AI Prompts: Load an Excel workbook with Aspose.Cells, assign CultureInfo('de-DE') to workbook.Settings, then retrieve the raw double from a cell and convert it to a culture‑aware string. | Demonstrate how to obtain the displayed text of a cell that uses locale‑specific number formatting by accessing Cell.StringValue after setting the workbook culture. | Show how to format a numeric value according to the workbook’s CultureInfo without altering the underlying double value.
// Common Searches: Aspose.Cells C# read cell value with German number format | How to set workbook culture to de-DE for numeric parsing in Aspose.Cells | Get displayed text of a formatted Excel cell using Aspose.Cells .NET | Retrieve raw double and formatted string from Excel cell with locale settings | Read locale‑specific numeric cells in Aspose.Cells without losing precision
// Tags: culture-aware numeric extraction Aspose.Cells | set workbook CultureInfo de-DE | read cell StringValue with locale formatting | format double using workbook settings | load Excel workbook with locale-specific number format

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel file with Aspose.Cells, sets the workbook's CultureInfo to German (de-DE), accesses a specific cell, obtains its raw double value, formats it using the workbook's culture, retrieves the displayed text via StringValue, and prints both the raw and culture‑aware representations.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    return;
                }

                // Load the workbook from the file
                Workbook workbook = new Workbook(inputPath);

                // Set the culture that matches the number formatting used in the worksheet (German example)
                workbook.Settings.CultureInfo = new CultureInfo("de-DE");

                // Access the first worksheet (index 0)
                Worksheet sheet = workbook.Worksheets[0];

                // Read a specific cell (e.g., B2) that contains a number formatted with the culture‑specific pattern
                Cell cell = sheet.Cells["B2"];

                // Retrieve the raw numeric value (as double) – culture does not affect the underlying value
                double rawNumber = cell.Value is double d ? d : 0.0;

                // Convert the numeric value to a string using the workbook's culture settings
                string formattedNumber = rawNumber.ToString(workbook.Settings.CultureInfo);

                // Get the displayed text as it appears in Excel (using StringValue which reflects the displayed format)
                string displayedText = cell.StringValue;

                // Output results
                Console.WriteLine($"Raw value: {rawNumber}");
                Console.WriteLine($"Formatted (culture aware): {formattedNumber}");
                Console.WriteLine($"Displayed text from cell: {displayedText}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
