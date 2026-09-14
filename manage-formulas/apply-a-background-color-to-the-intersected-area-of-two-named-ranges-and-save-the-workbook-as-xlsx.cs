// Title: Apply a solid yellow fill to the intersected cells of two ranges and save as XLSX using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates two cell ranges, finds their intersection, applies a solid yellow background to the intersected cells, and saves the workbook as an .xlsx file with Aspose.Cells. | Update the sample to define the ranges as named ranges and then shade the overlapping area with a custom color using Aspose.Cells. | Extend the program to copy the intersected range to a new worksheet while preserving the applied background style, then save the workbook.
// Common Searches: Aspose.Cells C# highlight cells where two ranges overlap | How to use Range.Intersect to apply background color in Aspose.Cells .NET | Saving a workbook with styled intersected range as XLSX using Aspose.Cells | C# example for applying solid fill to intersected range in Aspose.Cells
// Tags: range intersect background fill Aspose.Cells | cell shading with StyleFlag Aspose.Cells | save workbook as xlsx with styles Aspose.Cells | create and intersect ranges Aspose.Cells | apply solid color to overlapping cells .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// // This example creates a workbook, defines two ranges (B2:D5 and C4:E7), obtains their intersected area, applies a solid yellow fill to those cells using a Style and StyleFlag, and saves the result as IntersectedBackground.xlsx.
class ApplyIntersectionBackground
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define two ranges directly (no need for named ranges)
            Aspose.Cells.Range rangeA = sheet.Cells.CreateRange("B2:D5");
            Aspose.Cells.Range rangeB = sheet.Cells.CreateRange("C4:E7");

            // Get the intersected area of the two ranges
            Aspose.Cells.Range intersected = rangeA.Intersect(rangeB);
            if (intersected != null)
            {
                // Create a style with a solid yellow background
                Style style = workbook.CreateStyle();
                style.ForegroundColor = Color.Yellow;
                style.Pattern = BackgroundType.Solid;

                // Apply only cell shading
                StyleFlag flag = new StyleFlag { CellShading = true };
                intersected.ApplyStyle(style, flag);
            }

            // Define output path and ensure the directory exists
            string outputPath = "IntersectedBackground.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            // If outputDir is null (e.g., when only a file name is provided), use the current directory
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
                outputPath = Path.Combine(outputDir, outputPath);
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
