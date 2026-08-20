// Title: Aspose.Cells .NET: Export to TXT with Non‑Exponential Numbers Using TxtSaveOptions
// Description: Learn how to create a workbook, apply a fixed‑point style, set TxtSaveOptions.FormatStrategy to DisplayStyle, and save a tab‑separated TXT where all numbers appear in plain decimal (no scientific notation).
// Keywords: Aspose.Cells TxtSaveOptions | DisplayStyle CellValueFormatStrategy | prevent scientific notation | .NET export to TXT | plain decimal number format | tab‑separated values Aspose | Excel to text without exponential
// Common Searches: Aspose.Cells save as txt without scientific notation | TxtSaveOptions DisplayStyle example .NET | how to force plain decimal when exporting Excel | non‑exponential numbers in Aspose.Cells export | set number format for range before txt export
// Developer Intent: The developer needs to export an Excel workbook to a TXT file while guaranteeing that every numeric cell is rendered in fixed‑point (plain decimal) form rather than scientific notation.
// Use Cases: Generating TSV reports for downstream systems that cannot parse scientific notation. | Creating data files for legacy financial applications that require fixed‑point numbers. | Producing human‑readable text exports where large or tiny values must stay in decimal format for clarity.
// AI Prompts: Show me a C# example that uses TxtSaveOptions with CellValueFormatStrategy.DisplayStyle to export a workbook without scientific notation. | How can I apply a two‑decimal number style to a range and then save the sheet as a tab‑separated TXT in Aspose.Cells? | Explain the steps to configure SaveOptions so all numeric cells are saved as plain decimal values in a TXT file.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsNonExponentialDemo
{
    // Learn how to create a workbook, apply a fixed‑point style, set TxtSaveOptions.FormatStrategy to DisplayStyle, and save a tab‑separated TXT where all numbers appear in plain decimal (no scientific notation).
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate cells with numeric values that could be displayed in exponential form
                sheet.Cells["A1"].PutValue(123456789);
                sheet.Cells["A2"].PutValue(0.00000123);
                sheet.Cells["A3"].PutValue(1.23e+10);
                sheet.Cells["A4"].PutValue(3.14159265358979);

                // Apply a number format that forces plain decimal representation
                Style numberStyle = workbook.CreateStyle();
                numberStyle.Number = 2; // Two decimal places, no scientific notation

                // Set the style for the range A1:A4
                AsposeRange range = sheet.Cells.CreateRange("A1:A4");
                range.SetStyle(numberStyle);

                // Configure TxtSaveOptions to use the display style (formatted values) when exporting
                TxtSaveOptions txtOptions = new TxtSaveOptions
                {
                    Separator = '\t', // Tab‑separated values
                    FormatStrategy = CellValueFormatStrategy.DisplayStyle // Use formatted (non‑exponential) values
                };

                // Save the workbook to a TXT file using the configured options
                workbook.Save("NonExponentialOutput.txt", txtOptions);

                Console.WriteLine("Workbook saved with non‑exponential numeric display.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
