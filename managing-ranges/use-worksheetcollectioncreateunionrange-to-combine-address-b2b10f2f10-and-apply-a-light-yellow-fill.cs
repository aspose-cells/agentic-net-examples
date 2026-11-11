// Title: Combine non‑contiguous cells B2:B10 and F2:F10 into a union range and fill them with light yellow using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create union range for B2:B10 and F2:F10 | how to apply background color to non‑adjacent cells with Aspose.Cells | using WorksheetCollection.CreateUnionRange to set a fill style in .NET | Aspose.Cells example applying light yellow fill to a union range | apply StyleFlag CellShading to multiple address blocks in C#
// Tags: WorksheetCollection.CreateUnionRange with multiple address blocks | apply solid fill to union range Aspose.Cells | StyleFlag CellShading for noncontiguous cells | C# Aspose.Cells light yellow background | saving union range to .xlsx using Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The sample creates a new workbook, builds a union range covering B2:B10 and F2:F10, defines a style with a light‑yellow solid fill, applies the style to the union range using a StyleFlag, and saves the result as UnionRangeYellow.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (default sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Combine the two address blocks into a union range (sheet index 0)
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("B2:B10,F2:F10", 0);

            // Define a style with a light yellow solid fill
            Style yellowStyle = workbook.CreateStyle();
            yellowStyle.ForegroundColor = Color.LightYellow;
            yellowStyle.Pattern = BackgroundType.Solid;

            // Apply the fill style to the entire union range
            unionRange.ApplyStyle(yellowStyle, new StyleFlag() { CellShading = true });

            // Save the workbook to a file
            string outputPath = "UnionRangeYellow.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
