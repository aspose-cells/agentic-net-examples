// Title: Copy a Range with Values and Formatting to a New Workbook using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a source workbook, style a range (A1:C3), and copy both the cell values and the applied style to a matching range in a new workbook using Aspose.Cells' CopyData and CopyStyle methods, then save the result as CopiedRange.xlsx.
// Keywords: Aspose.Cells copy range C# | CopyData Aspose.Cells | CopyStyle Aspose.Cells | preserve formatting when copying cells | .NET Excel range transfer | copy styled cells between workbooks | Aspose.Cells example copy range with style
// Common Searches: Aspose.Cells copy range with formatting | How to preserve cell style when copying between workbooks in C# | CopyData vs CopyStyle Aspose.Cells | Copy styled range to new workbook Aspose.Cells .NET | Transfer Excel range values and styles programmatically
// Developer Intent: Copy a defined cell range from one workbook to another while keeping both the data and the original formatting intact.
// Use Cases: Generate a formatted report section in a separate workbook for client delivery. | Build a template workbook by reusing styled data blocks from an existing file. | Migrate specific styled tables during a data‑migration project without losing appearance.
// AI Prompts: Write C# code that copies the range A1:D5 from a source workbook to a destination workbook using Aspose.Cells, preserving values and all formatting. | Explain when to use CopyData, CopyStyle, or the combined Copy method in Aspose.Cells and the impact on performance. | Provide a step‑by‑step guide for copying multiple non‑contiguous ranges with their styles into a new workbook using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a source workbook, style a range (A1:C3), and copy both the cell values and the applied style to a matching range in a new workbook using Aspose.Cells' CopyData and CopyStyle methods, then save the result as CopiedRange.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create source workbook and populate a range ----------
            Workbook srcWorkbook = new Workbook();
            Worksheet srcSheet = srcWorkbook.Worksheets[0];
            Cells srcCells = srcSheet.Cells;

            // Fill A1:C3 with sample values
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    srcCells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Apply a style to the source range
            Style srcStyle = srcWorkbook.CreateStyle();
            srcStyle.Font.Name = "Arial";
            srcStyle.Font.Size = 12;
            srcStyle.Font.IsBold = true;
            srcStyle.ForegroundColor = Color.LightBlue;
            srcStyle.Pattern = BackgroundType.Solid;
            srcCells.CreateRange("A1:C3").SetStyle(srcStyle);

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // Define source and destination ranges using the Aspose.Cells.Range alias
            AsposeRange sourceRange = srcCells.CreateRange("A1:C3");
            AsposeRange destinationRange = destCells.CreateRange("A1:C3");

            // Copy cell values from source to destination
            destinationRange.CopyData(sourceRange);

            // Copy formatting (style) from source to destination
            destinationRange.CopyStyle(sourceRange);

            // ---------- Save the destination workbook ----------
            string outputPath = "CopiedRange.xlsx";
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
