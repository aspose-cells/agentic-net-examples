// Title: Aspose.Cells for .NET – C# Example: Copy formatting from a range offset five rows upward
// Description: This C# demo creates a workbook, defines an original range (A6:D10) and an offset range five rows above it (A1:D5), applies a custom style to the offset range, copies only the formatting back to the original range with CopyStyle, optionally adds sample values, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | .NET | CopyStyle | offset range | copy formatting | Excel cell style | range manipulation | example code | Aspose.Cells API | formatting only
// Common Searches: Aspose.Cells copy style from another range | C# offset range and copy formatting | Copy formatting of a range five rows above | Aspose.Cells CopyStyle method example | how to copy cell styles without values in .NET
// Developer Intent: Copy only the formatting of a source range located five rows above a target range and apply it to the target range using Aspose.Cells.
// Use Cases: Reuse a header style defined at the top of a sheet for data blocks placed further down. | Apply a hidden template's formatting to multiple report sections after inserting rows. | Synchronize cell appearance after moving data by copying style from a reference range. | Create a quick style‑swap utility that updates formatting based on a predefined offset range.
// AI Prompts: Generate C# code that uses Aspose.Cells to copy only the formatting from a source range offset by N rows to a destination range of the same size. | Show how to apply the CopyStyle method to transfer cell styles without values between two equal‑sized ranges in Aspose.Cells for .NET. | Explain how to calculate an offset range based on an original range's dimensions and then copy its style back to the original range.

using System;
using Aspose.Cells;
using System.Drawing;

// This C# demo creates a workbook, defines an original range (A6:D10) and an offset range five rows above it (A1:D5), applies a custom style to the offset range, copies only the formatting back to the original range with CopyStyle, optionally adds sample values, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the original range (the location where formatting will be copied back to)
            // Example: cells A6:D10
            Aspose.Cells.Range originalRange = cells.CreateRange("A6:D10");

            // Define the offset range which is 5 rows above the original range
            // Example: cells A1:D5
            Aspose.Cells.Range offsetRange = cells.CreateRange("A1:D5");

            // ------------------------------------------------------------
            // Create a sample style and apply it to the offset range
            // ------------------------------------------------------------
            Style style = workbook.CreateStyle();
            style.Font.Name = "Arial";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.Blue;
            style.ForegroundColor = Color.LightYellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to every cell in the offset range
            for (int row = 0; row < offsetRange.RowCount; row++)
            {
                for (int col = 0; col < offsetRange.ColumnCount; col++)
                {
                    offsetRange[row, col].SetStyle(style);
                    // Add some dummy values so the cells are visible
                    offsetRange[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ------------------------------------------------------------
            // Copy the formatting from the offset range back to the original range
            // ------------------------------------------------------------
            originalRange.CopyStyle(offsetRange);

            // (Optional) Put values in the original range to see the effect
            for (int row = 0; row < originalRange.RowCount; row++)
            {
                for (int col = 0; col < originalRange.ColumnCount; col++)
                {
                    originalRange[row, col].PutValue($"Orig {row + 1},{col + 1}");
                }
            }

            // Save the workbook
            workbook.Save("OffsetCopyFormattingDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
