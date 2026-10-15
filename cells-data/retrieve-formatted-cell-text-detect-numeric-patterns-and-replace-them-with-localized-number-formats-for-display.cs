// Title: Localize numeric substrings in Excel cell text using Aspose.Cells and C# (e.g., German number format)
// AI Prompts: Generate C# code with Aspose.Cells that iterates over all used cells, extracts the displayed string, finds numeric patterns with a regular expression, and rewrites each number using a given CultureInfo. | Demonstrate how to change decimal separators from '.' to ',' only in the numeric parts of mixed‑content Excel cells while leaving surrounding text untouched, using Aspose.Cells. | Create a reusable method that parses invariant numbers inside cell text, formats them with a target locale, assigns the result back as a string, clears existing number formats, and saves the workbook.
// Common Searches: aspnet aspose.cells replace decimal separator in cell text german locale | c# regex find numbers in Excel cell string and format with cultureinfo using aspose | how to localize numbers inside mixed text cells in Excel with Aspose.Cells .NET | convert invariant numeric values to locale‑specific format in Aspose.Cells worksheet | apply German number formatting to numbers embedded in Excel cell strings c#
// Tags: Aspose.Cells replace numbers in cell strings | C# regex numeric detection in Excel cells | culture-specific number formatting with Aspose.Cells | localize decimal separator in Excel worksheet | update cell value as string Aspose.Cells

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, walks through each used cell, reads the cell's displayed text, uses a regular expression to locate numeric substrings, parses them with invariant culture, formats them according to a specified CultureInfo (e.g., German), replaces the original numbers while preserving surrounding text, writes the result back as a plain string, clears any existing number format to avoid double formatting, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Choose the worksheet to process (first worksheet in this example)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define a regex to detect numeric patterns (integers and decimals, optional sign)
        Regex numberRegex = new Regex(@"-?\d+(\.\d+)?", RegexOptions.Compiled);

        // Define the culture you want to use for localized formatting
        // For example, German (Germany) uses comma as decimal separator
        CultureInfo targetCulture = new CultureInfo("de-DE");

        // Iterate through all used cells in the worksheet
        foreach (Cell cell in cells)
        {
            // Retrieve the cell's formatted (display) text
            string formattedText = cell.StringValue;

            // If the cell is empty, skip processing
            if (string.IsNullOrEmpty(formattedText))
                continue;

            // Replace each numeric occurrence with its localized representation
            string localizedText = numberRegex.Replace(formattedText, match =>
            {
                // Try to parse the matched number using invariant culture (dot as decimal separator)
                if (double.TryParse(match.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
                {
                    // Format the number using the target culture's number format (e.g., "N" for numeric)
                    return number.ToString("N", targetCulture);
                }
                // If parsing fails, return the original match unchanged
                return match.Value;
            });

            // If any changes were made, update the cell's value with the new localized string
            if (!localizedText.Equals(formattedText, StringComparison.Ordinal))
            {
                // Set the cell's value as a string to preserve the exact display text
                cell.PutValue(localizedText);
                // Optionally, you can clear any number format to avoid double formatting
                cell.GetStyle().Number = 0;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
