// Title: Apply a light‑yellow background to a non‑contiguous union range (B2:B10,F2:F10) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply light yellow fill to multiple non‑adjacent ranges | How to use WorksheetCollection.CreateUnionRange to set cell shading in .NET | C# example for creating a union range and applying a solid background with Aspose.Cells | Set background color for B2:B10 and F2:F10 together using Aspose.Cells | Apply cell shading to a union range without affecting other formatting Aspose.Cells
// Tags: CreateUnionRange background fill Aspose.Cells | union range cell shading .NET | noncontiguous range fill C# | light yellow solid pattern Aspose.Cells | WorksheetCollection union range example

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a workbook, defines a union range that combines cells B2:B10 and F2:F10 on the first worksheet, builds a light‑yellow solid fill style, applies only the cell shading to that union range, and saves the result as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook with a default worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Create a union range that combines "B2:B10" and "F2:F10" on the first sheet (index 0)
                UnionRange unionRange = workbook.Worksheets.CreateUnionRange("B2:B10,F2:F10", 0);

                // Define a light‑yellow fill style
                Style fillStyle = workbook.CreateStyle();
                fillStyle.ForegroundColor = Color.LightYellow;
                fillStyle.Pattern = BackgroundType.Solid;

                // Apply only the fill (cell shading) to the union range
                StyleFlag flag = new StyleFlag { CellShading = true };
                unionRange.ApplyStyle(fillStyle, flag);

                // Save the workbook
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
