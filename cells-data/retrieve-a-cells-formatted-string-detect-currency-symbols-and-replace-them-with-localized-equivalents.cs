// Title: Read a cell's formatted string and replace currency symbols with locale codes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get displayed value of a formatted cell | replace $ € £ ¥ with USD EUR GBP JPY in Excel using Aspose.Cells | how to map currency symbols to locale codes in Aspose.Cells .NET | read formatted cell value and write localized string to another cell with Aspose.Cells | use regex to substitute currency symbols in an Excel cell via Aspose.Cells
// Tags: formatted cell value retrieval Aspose.Cells | currency symbol to locale code mapping .NET | regex based currency replacement Excel | write localized string to another cell Aspose.Cells | apply built-in currency number format Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example creates a workbook (if missing), applies a built‑in currency number format to cell B2, reads the cell's displayed (formatted) string, uses a regular expression to replace common currency symbols ($, €, £, ¥) with their locale codes (USD, EUR, GBP, JPY), writes the localized text to cell C2, and saves the modified workbook.
class CurrencyLocalization
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                // Sample value with currency format
                ws.Cells["B2"].PutValue(1234.56);
                // Apply a currency number format (e.g., $#,##0.00)
                var style = ws.Cells["B2"].GetStyle();
                style.Number = 164;
                ws.Cells["B2"].SetStyle(style);
                wb.Save(inputPath);
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Target cell whose formatted string we want to process
            var cell = worksheet.Cells["B2"];

            // Retrieve the cell's displayed (formatted) string
            string formatted = cell.StringValue;

            // Mapping of common currency symbols to their localized textual equivalents
            var currencyMap = new Dictionary<string, string>
            {
                { "$", "USD" },
                { "€", "EUR" },
                { "£", "GBP" },
                { "¥", "JPY" }
            };

            // Replace any currency symbol found in the formatted string with the mapped value
            string localized = Regex.Replace(formatted, @"[\$\€\£\¥]", match =>
            {
                string symbol = match.Value;
                return currencyMap.TryGetValue(symbol, out var replacement) ? replacement : symbol;
            });

            // Display the original and the localized strings
            Console.WriteLine("Original formatted: " + formatted);
            Console.WriteLine("Localized string   : " + localized);

            // Write the localized string back to another cell
            worksheet.Cells["C2"].PutValue(localized);

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
