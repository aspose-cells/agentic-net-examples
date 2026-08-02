// Title: Copy only formatting from a source range to a destination range using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create two worksheets, apply a custom style to a source range (A1:C3), and transfer that style to a destination range on another sheet without changing the destination's cell values. The example uses the Range.CopyStyle method and saves the result as an Excel file.
// Keywords: Aspose.Cells C# copy formatting | Range.CopyStyle .NET | copy style without values Aspose | preserve cell data while copying style | Excel formatting transfer Aspose.Cells | C# Aspose.Cells range style example
// Common Searches: Aspose.Cells copy only formatting C# | Range.CopyStyle example .NET | how to copy style between ranges without data loss | preserve cell values when copying Excel style with Aspose | copy formatting from one worksheet to another Aspose.Cells
// Developer Intent: Transfer the visual style of a source range to another range while keeping the target cells' existing values intact.
// Use Cases: Apply a corporate header style from a template sheet to multiple data sheets without overwriting the data. | Synchronize font, color, and border settings across worksheets generated in a reporting pipeline. | Migrate a programmatically created custom style to a range in a different workbook while preserving its content.
// AI Prompts: Show C# code that copies only the formatting of a range to another range using Aspose.Cells, leaving the destination values unchanged. | Explain how to use Range.CopyStyle in Aspose.Cells for .NET to transfer styles between worksheets without affecting cell data. | Provide a step‑by‑step example of creating a custom style, applying it to a source range, and copying that style to a destination range in Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create two worksheets, apply a custom style to a source range (A1:C3), and transfer that style to a destination range on another sheet without changing the destination's cell values. The example uses the Range.CopyStyle method and saves the result as an Excel file.
class CopyFormattingDemo
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

            // Fill source range with data
            Cells srcCells = srcSheet.Cells;
            AsposeRange srcRange = srcCells.CreateRange("A1:C3");
            srcRange[0, 0].PutValue("A1");
            srcRange[0, 1].PutValue("B1");
            srcRange[0, 2].PutValue("C1");
            srcRange[1, 0].PutValue("A2");
            srcRange[1, 1].PutValue("B2");
            srcRange[1, 2].PutValue("C2");
            srcRange[2, 0].PutValue("A3");
            srcRange[2, 1].PutValue("B3");
            srcRange[2, 2].PutValue("C3");

            // Create a style and apply it to the source range
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;
            srcRange.SetStyle(style);

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            destSheet.Name = "Destination";

            // Create destination range with its own values
            Cells destCells = destSheet.Cells;
            AsposeRange destRange = destCells.CreateRange("A1:C3");
            destRange[0, 0].PutValue("X1");
            destRange[0, 1].PutValue("Y1");
            destRange[0, 2].PutValue("Z1");
            destRange[1, 0].PutValue("X2");
            destRange[1, 1].PutValue("Y2");
            destRange[1, 2].PutValue("Z2");
            destRange[2, 0].PutValue("X3");
            destRange[2, 1].PutValue("Y3");
            destRange[2, 2].PutValue("Z3");

            // Copy only the formatting from source range to destination range
            destRange.CopyStyle(srcRange);

            // Save the workbook
            workbook.Save("CopyFormattingDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
