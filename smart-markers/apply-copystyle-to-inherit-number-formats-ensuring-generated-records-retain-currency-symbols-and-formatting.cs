// Title: CopyStyle Preserves Currency Number Format in Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create a built‑in currency style in a source worksheet, apply it to a range, and then copy the style to a destination range using Aspose.Cells CopyStyle. The destination cells retain the currency symbols and number formatting after values are written, making it ideal for financial Excel reports.
// Keywords: Aspose.Cells | CopyStyle | currency number format | C# | .NET | Excel style copy | preserve number format | financial report generation | range styling
// Common Searches: Aspose.Cells copy style with currency format | CopyStyle retain number format C# | How to copy built‑in number format using Aspose.Cells | C# copy Excel cell style preserving currency symbols | Aspose.Cells range style inheritance
// Developer Intent: Copy a source range’s style so that the destination range keeps the same currency number format.
// Use Cases: Generating multi‑sheet financial statements where currency formatting must stay consistent. | Duplicating styled templates for invoices or budgets across worksheets. | Automating Excel exports that require exact replication of number formats such as currency, percentages, or dates.
// AI Prompts: Show me a C# Aspose.Cells example that copies a currency number format from one range to another using CopyStyle. | Explain how CopyStyle preserves built‑in number formats like currency when transferring styles between worksheets. | Provide code that verifies the currency symbols remain after copying styles with Aspose.Cells.

using System;
using Aspose.Cells;

// This example demonstrates how to create a built‑in currency style in a source worksheet, apply it to a range, and then copy the style to a destination range using Aspose.Cells CopyStyle. The destination cells retain the currency symbols and number formatting after values are written, making it ideal for financial Excel reports.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet ----------
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Create a style with a built‑in currency number format
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 5;                     // Built‑in currency format
            currencyStyle.IsNumberFormatApplied = true;  // Ensure the format is applied

            // Apply the style to a source range and put numeric values
            Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:A3");
            srcRange.SetStyle(currencyStyle);
            srcRange[0, 0].PutValue(1234.56);
            srcRange[1, 0].PutValue(7890);
            srcRange[2, 0].PutValue(-45.67);

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            destSheet.Name = "Destination";

            // Create a destination range of the same size
            Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("B1:B3");

            // Copy the style (including number format) from the source range
            destRange.CopyStyle(srcRange);

            // Put the same numeric values to verify the currency formatting is retained
            destRange[0, 0].PutValue(1234.56);
            destRange[1, 0].PutValue(7890);
            destRange[2, 0].PutValue(-45.67);

            // Save the workbook
            workbook.Save("CopyStyleCurrencyDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
