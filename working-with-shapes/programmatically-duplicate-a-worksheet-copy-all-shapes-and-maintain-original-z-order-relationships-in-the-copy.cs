// Title: How to duplicate an Excel worksheet and copy all shapes while preserving Z‑order using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that duplicates a worksheet, clears the automatically copied shapes, and re‑adds each shape with its original cell range and ZOrderPosition using Aspose.Cells. | Create a console program that loads a workbook, copies Sheet1 to a new sheet, iterates over sourceSheet.Shapes, adds each shape to the copy with AddCopy, and sets newShape.ZOrderPosition to match the source shape.
// Common Searches: asp.net copy worksheet with shapes preserving layer order using Aspose.Cells | c# aspose.cells duplicate sheet and keep shape Z-order | how to retain shape stacking order when copying an Excel sheet with Aspose.Cells | copy Excel worksheet and preserve shape positions and ZOrderPosition in .NET
// Tags: worksheet duplication with shape ZOrder Aspose.Cells | copy Excel sheet shapes preserving layering C# | Aspose.Cells AddCopy shape ZOrderPosition | clear auto‑copied shapes before re‑adding Aspose.Cells | preserve shape order when cloning worksheet

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads input.xlsx, creates a copy of Sheet1 named Sheet1_Copy, removes the shapes that Aspose.Cells copies automatically, then iterates through each shape on the original sheet, adds a copy to the new sheet with the same cell coordinates using AddCopy, and assigns the original ZOrderPosition to maintain the exact layering before saving to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the worksheet to duplicate
            Worksheet sourceSheet = workbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            int sourceIndex = sourceSheet.Index;

            // Add a copy of the worksheet
            int copyIndex = workbook.Worksheets.AddCopy(sourceIndex);
            Worksheet copySheet = workbook.Worksheets[copyIndex];
            copySheet.Name = sourceSheet.Name + "_Copy";

            // Preserve original Z‑order of shapes
            // Clear any automatically copied shapes
            copySheet.Shapes.Clear();

            // Copy each shape and retain its ZOrderPosition
            foreach (Shape srcShape in sourceSheet.Shapes)
            {
                try
                {
                    // Copy shape with its original placement
                    Shape newShape = copySheet.Shapes.AddCopy(
                        srcShape,
                        srcShape.UpperLeftRow,
                        srcShape.UpperLeftColumn,
                        srcShape.LowerRightRow,
                        srcShape.LowerRightColumn);

                    // Preserve Z‑order
                    newShape.ZOrderPosition = srcShape.ZOrderPosition;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to copy shape '{srcShape.Name}': {ex.Message}");
                }
            }

            // Save the workbook with the duplicated worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
