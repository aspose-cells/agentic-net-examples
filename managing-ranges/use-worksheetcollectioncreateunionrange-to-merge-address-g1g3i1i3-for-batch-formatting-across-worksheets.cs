// Title: Apply a yellow fill to a union range G1:G3,I1:I3 across all worksheets using WorksheetCollection.CreateUnionRange in Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a union range for cells G1:G3 and I1:I3 on every worksheet in a workbook and applies a solid yellow background using Aspose.Cells. | Demonstrate how to use WorksheetCollection.CreateUnionRange to batch‑style the same cell addresses on multiple sheets in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells how to format the same cell range on all worksheets in C# | CreateUnionRange example for applying style to multiple sheets | Batch apply yellow fill to G1:G3 and I1:I3 across workbook using Aspose.Cells | WorksheetCollection.CreateUnionRange usage for uniform cell styling | C# Aspose.Cells union range across worksheets tutorial
// Tags: union range styling Aspose.Cells | WorksheetCollection.CreateUnionRange C# | apply solid fill to multiple sheets Aspose.Cells | batch cell formatting across worksheets | G1:G3 I1:I3 union range Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// The sample creates a workbook with three worksheets, builds a union range that includes cells G1:G3 and I1:I3 on every sheet via WorksheetCollection.CreateUnionRange, applies a solid yellow fill style to the range, and saves the file as UnionRangeFormatted.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Add additional worksheets
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Create a union range that spans the same address on every worksheet in the collection,
            // starting from the first worksheet (index 0)
            UnionRange unionRange = workbook.Worksheets.CreateUnionRange("G1:G3,I1:I3", 0);

            // Define a style (yellow fill)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.Yellow;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the union range
            unionRange.ApplyStyle(style, new StyleFlag { All = true });

            // Save the workbook
            string outputPath = "UnionRangeFormatted.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
