// Title: Create a UnionRange for cells B2:B10 and F2:F10 and apply a light‑yellow fill with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses WorksheetCollection.CreateUnionRange to merge the ranges B2:B10 and F2:F10, then applies a solid light‑yellow background style to the resulting UnionRange. | Show how to define a Style with a light‑yellow foreground color and a solid pattern, and apply it only to cell shading of a UnionRange created in Aspose.Cells.
// Common Searches: Aspose.Cells C# create union range for non‑contiguous cells and set fill color | How to apply a solid background to multiple separate ranges using WorksheetCollection.CreateUnionRange | Example of styling a UnionRange with yellow fill in Aspose.Cells for .NET | C# code to combine B2:B10 and F2:F10 into one range and change cell shading
// Tags: WorksheetCollection.CreateUnionRange C# example | apply solid fill to UnionRange Aspose.Cells | yellow fill style Aspose.Cells | noncontiguous cell range styling .NET | UnionRange shading C#

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// // This program creates a new workbook, builds a UnionRange that combines cells B2:B10 and F2:F10 on the first worksheet, defines a light‑yellow solid fill style, applies the style to the UnionRange, and saves the workbook as UnionRangeYellowFill.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Create a union range that combines the two address blocks.
            // WorksheetCollection.CreateUnionRange returns a UnionRange object.
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("B2:B10,F2:F10", 0);

            // Prepare a style with a light yellow fill
            Style yellowStyle = workbook.CreateStyle();
            yellowStyle.ForegroundColor = Color.LightYellow;
            yellowStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the union range (only cell shading)
            StyleFlag flag = new StyleFlag { CellShading = true };
            unionRange.ApplyStyle(yellowStyle, flag);

            // Define output file path
            string outputPath = "UnionRangeYellowFill.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
