// Title: How to read a cell's displayed string and replace currency symbols with ISO codes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get cell display text with formatting | Replace currency symbols with ISO codes in an Excel workbook using Aspose.Cells | Use Regex to map $ € £ ¥ to locale abbreviations in Aspose.Cells | Read formatted cell value and write localized string to another cell in C# | Map Excel currency symbols to locale codes with Aspose.Cells API
// Tags: Cell.GetStringValue display string Aspose.Cells | currency symbol to ISO code mapping C# | regex replace cell content Aspose.Cells | write localized string to Excel cell Aspose.Cells | load and save workbook Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an existing workbook, retrieves the formatted display string of cell A1 using GetStringValue with the DisplayString strategy, replaces currency symbols ($, €, £, ¥) with their ISO code equivalents (USD, EUR, GBP, JPY) via a regular expression and a dictionary, writes the localized result to cell B1, and saves the workbook as a new file.
class CurrencyLocalization
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Mapping of currency symbols to localized equivalents
            var currencyMap = new Dictionary<string, string>()
            {
                { "$", "USD" },
                { "€", "EUR" },
                { "£", "GBP" },
                { "¥", "JPY" }
                // Add additional mappings as needed
            };

            // Example: process cell A1
            Cell cell = worksheet.Cells["A1"];

            // Retrieve the cell's formatted string as displayed in Excel
            string formattedString = cell.GetStringValue(CellValueFormatStrategy.DisplayString);

            // Replace currency symbols with their localized equivalents
            string localizedString = Regex.Replace(formattedString, @"[\$\€\£\¥]", match =>
            {
                string symbol = match.Value;
                return currencyMap.ContainsKey(symbol) ? currencyMap[symbol] : symbol;
            });

            // Write the localized string to another cell (e.g., B1)
            worksheet.Cells["B1"].PutValue(localizedString);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
