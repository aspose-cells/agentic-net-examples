// Title: Aspose.Cells .NET: Create a Named Union Range of Non‑Contiguous Cells and Apply a Custom Style (C#)
// Description: This example shows how to build a UnionRange that combines A1:A5 and C1:C5, assign it the name "MyUnionRange", define a solid light‑green background with bold dark‑blue centered text, apply the style to the whole range, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells UnionRange C# | named range non‑contiguous cells | apply custom style Aspose.Cells | Union method Aspose.Cells .NET | StyleFlag all properties | C# Excel formatting Aspose
// Common Searches: how to create a named union range in Aspose.Cells .NET | apply a solid background and bold font to a union range C# | Aspose.Cells UnionRanges example with custom styling | C# code to style non‑adjacent cells in Excel using Aspose
// Developer Intent: Generate a named union range that spans separate cell blocks and format it with a custom style in a .NET workbook.
// Use Cases: Group columns A and C for a single data‑validation rule or chart series. | Highlight scattered sections of a report with consistent formatting. | Reference a multi‑area range in formulas to compute totals across non‑adjacent cells.
// AI Prompts: Write C# code with Aspose.Cells to create a union range covering A1:A5 and C1:C5, name it "MyUnionRange", and apply a light‑green background with bold dark‑blue centered text. | Explain how to modify only the font color of an existing UnionRange in Aspose.Cells without changing its background or alignment. | Provide step‑by‑step instructions for using UnionRanges and ApplyStyle with a StyleFlag to style non‑contiguous cells in an Excel workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsUnionRangeExample
{
    // This example shows how to build a UnionRange that combines A1:A5 and C1:C5, assign it the name "MyUnionRange", define a solid light‑green background with bold dark‑blue centered text, apply the style to the whole range, and save the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create the first contiguous range (A1:A5)
                AsposeRange range1 = worksheet.Cells.CreateRange("A1:A5");

                // Convert the first range to a UnionRange
                UnionRange unionRange = range1.UnionRanges(new AsposeRange[] { range1 });

                // Add a second non‑contiguous range (C1:C5) using the Union method
                unionRange = unionRange.Union("C1:C5");

                // Assign a name to the union range
                unionRange.Name = "MyUnionRange";

                // Create a custom style
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightGreen;
                style.Font.IsBold = true;
                style.Font.Color = Color.DarkBlue;
                style.HorizontalAlignment = TextAlignmentType.Center;
                style.VerticalAlignment = TextAlignmentType.Center;

                // Apply the style to the entire union range
                unionRange.ApplyStyle(style, new StyleFlag { All = true });

                // Save the workbook
                string outputPath = "UnionRangeNamedStyle.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
