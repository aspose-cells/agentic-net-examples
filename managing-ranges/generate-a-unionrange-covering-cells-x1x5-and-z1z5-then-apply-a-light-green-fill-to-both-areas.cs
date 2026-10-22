// Title: Apply a light‑green background to non‑contiguous ranges X1:X5 and Z1:Z5 using Aspose.Cells for .NET
// Common Searches: how to shade non‑adjacent cells X1:X5 and Z1:Z5 with Aspose.Cells in C# | Aspose.Cells C# create union of X1:X5 and Z1:Z5 and set solid fill | apply solid light green style to multiple ranges in an Aspose.Cells workbook | C# Aspose.Cells apply cell shading to separate ranges without merging | set background color for disjoint cell blocks using Aspose.Cells .NET API
// Tags: apply solid fill to disjoint ranges Aspose.Cells | non‑adjacent range cell shading C# | union range creation Aspose.Cells .NET | light‑green background style Aspose.Cells | range X1:X5 Z1:Z5 Aspose.Cells example

using System;
using System.Collections;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, defines two separate ranges (X1:X5 and Z1:Z5) on the first worksheet, builds a union list, creates a solid light‑green fill style, applies the shading to each range, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Create two separate ranges: X1:X5 and Z1:Z5
            AsposeRange rangeX = sheet.Cells.CreateRange("X1:X5");
            AsposeRange rangeZ = sheet.Cells.CreateRange("Z1:Z5");

            // Combine the two ranges manually (UnionRanges is not a static method)
            ArrayList unionRanges = new ArrayList { rangeX, rangeZ };

            // Define a light‑green fill style
            Style greenStyle = workbook.CreateStyle();
            greenStyle.ForegroundColor = Color.LightGreen;
            greenStyle.Pattern = BackgroundType.Solid;

            // Apply the style to each range in the union (only cell shading)
            StyleFlag flag = new StyleFlag { CellShading = true };
            foreach (AsposeRange r in unionRanges)
            {
                r.ApplyStyle(greenStyle, flag);
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists (if a directory part is present)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
