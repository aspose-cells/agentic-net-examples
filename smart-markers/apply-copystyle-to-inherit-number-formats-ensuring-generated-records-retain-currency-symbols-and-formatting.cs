// Title: CopyStyle Preserves Currency Number Format in Aspose.Cells for .NET
// Description: Demonstrates how to apply a built‑in currency style (Number = 5) to a source range, copy the style with CopyStyle to another range, and retain the $ symbol and formatting after inserting values, using Aspose.Cells C# API.
// Keywords: Aspose.Cells | CopyStyle | currency number format | preserve number format | C# Excel library | built‑in number format | Excel range style copy | financial report automation | .NET Excel export
// Common Searches: Aspose.Cells CopyStyle keep currency symbol | CopyStyle retain number format .NET | how to copy built‑in number format with Aspose.Cells | C# copy Excel style including currency | preserve formatting when copying ranges Aspose
// Developer Intent: Copy a style from one cell range to another so the destination cells keep the exact currency number format applied to the source.
// Use Cases: Generate financial statements where calculated cells must display the same currency format as template cells. | Clone a formatted template block to a new area for dynamic data insertion while preserving monetary symbols. | Apply consistent number formatting across multiple worksheets in a multi‑sheet workbook.
// AI Prompts: Show how to use Aspose.Cells CopyStyle in C# to copy a built‑in currency format from range A1:B2 to C1:D2 and verify the $ symbol appears on the new cells. | Provide a C# example that copies styles, including number formats, across worksheets and saves the workbook with Aspose.Cells. | Explain why setting IsNumberFormatApplied = true is required when copying number formats with CopyStyle.

using System;
using Aspose.Cells;
using Aspose.Cells; // Ensure Aspose.Cells namespace is available
using System.Drawing;

// Demonstrates how to apply a built‑in currency style (Number = 5) to a source range, copy the style with CopyStyle to another range, and retain the $ symbol and formatting after inserting values, using Aspose.Cells C# API.
class CopyStyleNumberFormatDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill the source range with sample numeric values
            sheet.Cells["A1"].PutValue(1234.56);
            sheet.Cells["A2"].PutValue(7890.12);
            sheet.Cells["B1"].PutValue(345.67);
            sheet.Cells["B2"].PutValue(890.23);

            // Create a style that uses a built‑in currency format (includes the $ symbol)
            Style currencyStyle = workbook.CreateStyle();
            currencyStyle.Number = 5;                     // Built‑in currency format
            currencyStyle.IsNumberFormatApplied = true;  // Ensure the number format is applied

            // Apply the currency style to the source range A1:B2
            Aspose.Cells.Range srcRange = sheet.Cells.CreateRange("A1:B2");
            srcRange.SetStyle(currencyStyle);

            // Define the destination range C1:D2
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange("C1:D2");

            // Copy the style (including number format) from the source range to the destination range
            destRange.CopyStyle(srcRange);

            // Populate the destination range with values to verify the formatting is retained
            sheet.Cells["C1"].PutValue(5555.55);
            sheet.Cells["C2"].PutValue(6666.66);
            sheet.Cells["D1"].PutValue(7777.77);
            sheet.Cells["D2"].PutValue(8888.88);

            // Save the workbook
            workbook.Save("CopyStyleNumberFormatDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
