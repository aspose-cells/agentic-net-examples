// Title: Insert an SVG icon into a merged cell range (C3:D4) and make it scale proportionally with Aspose.Cells for .NET
// AI Prompts: Load an SVG file from a stream, add it as a picture to cell C3, set its Placement to MoveAndSize, and align its Left, Top, Width, and Height to the merged range C3:D4. | Create the merged range C3:D4, obtain its dimensions, and resize the inserted SVG picture so it fills the merged cells while preserving its aspect ratio.
// Common Searches: how to add an SVG image to merged cells in Aspose.Cells C# | scale picture to fit merged range C3:D4 using Aspose.Cells for .NET | Aspose.Cells MoveAndSize placement for SVG picture in Excel | C# insert SVG from file stream into Excel merged cells with Aspose
// Tags: insert SVG picture Aspose.Cells C# | merged cells picture scaling Aspose.Cells | picture placement MoveAndSize Aspose.Cells | load SVG from file stream Aspose.Cells | resize picture to merged range dimensions

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Range = Aspose.Cells.Range;

// The example creates a workbook, merges cells C3:D4, loads an SVG file, inserts it as a picture with MoveAndSize placement, aligns and resizes the picture to the merged range, and saves the workbook to output.xlsx.
class InsertSvgIntoMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge the range C3:D4 (zero‑based indices: rows 2‑3, columns 2‑3)
            sheet.Cells.Merge(2, 2, 2, 2);

            string svgPath = "icon.svg";
            if (!File.Exists(svgPath))
                throw new FileNotFoundException($"SVG file not found: {svgPath}");

            // Load the SVG icon from a file and insert it
            using (FileStream svgStream = File.OpenRead(svgPath))
            {
                // Add picture and obtain its index
                int pictureIndex = sheet.Pictures.Add(2, 2, svgStream);
                Picture pic = sheet.Pictures[pictureIndex];

                // Ensure the picture moves and resizes with the cells
                pic.Placement = PlacementType.MoveAndSize;

                // Get the merged range object for positioning and sizing
                Range mergedRange = sheet.Cells.CreateRange("C3:D4");

                // Align the picture to the merged range and scale it proportionally
                pic.Left = (int)mergedRange.Left;
                pic.Top = (int)mergedRange.Top;
                pic.Width = (int)mergedRange.Width;
                pic.Height = (int)mergedRange.Height;
            }

            // Save the workbook to a file
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
