// Title: How to insert an SVG icon into a merged cell range (C3:D4) and scale it proportionally using Aspose.Cells for .NET
// AI Prompts: Add an SVG picture to the merged range C3:D4, set Placement = MoveAndSize, and adjust its Width and Height to match the merged cells while keeping the original aspect ratio. | Calculate the combined pixel width and height of a merged cell area and apply those dimensions to a worksheet picture in C# with Aspose.Cells.
// Common Searches: Aspose.Cells .NET insert SVG into merged cells and keep aspect ratio | C# scale picture to fit merged cell range C3:D4 using Aspose.Cells | How to set picture placement MoveAndSize for merged cells in Aspose.Cells | Compute pixel dimensions of merged cells for image scaling Aspose.Cells C#
// Tags: insert SVG picture merged cells Aspose.Cells | picture placement MoveAndSize .NET | scale worksheet image to merged cell dimensions | calculate merged cell pixel size C# | preserve aspect ratio SVG Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, merges cells C3:D4, loads an SVG file, adds it as a picture positioned at the merged range, sets the picture's placement to MoveAndSize, computes the total pixel width and height of the merged cells, resizes the picture to those dimensions while preserving its aspect ratio, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells C3:D4 (rows 2‑3, columns 2‑3 in zero‑based indexing)
            sheet.Cells.Merge(2, 2, 2, 2); // merge 2 rows and 2 columns starting at C3

            // Define the merged range area
            CellArea mergedArea = new CellArea
            {
                StartRow = 2,
                StartColumn = 2,
                EndRow = 3,
                EndColumn = 3
            };

            // Path to the SVG image
            string svgPath = "icon.svg";

            // Ensure the SVG file exists before adding it
            if (!File.Exists(svgPath))
                throw new FileNotFoundException($"The SVG file '{svgPath}' was not found.");

            // Add the picture to the worksheet; Add returns the picture index
            int pictureIndex = sheet.Pictures.Add(2, 2, svgPath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Make the picture move and size with the cells
            picture.Placement = PlacementType.MoveAndSize;

            // Calculate total width of the merged cells in pixels
            double totalWidth = 0;
            for (int col = mergedArea.StartColumn; col <= mergedArea.EndColumn; col++)
                totalWidth += sheet.Cells.GetColumnWidthPixel(col);

            // Calculate total height of the merged cells in pixels
            double totalHeight = 0;
            for (int row = mergedArea.StartRow; row <= mergedArea.EndRow; row++)
                totalHeight += sheet.Cells.GetRowHeightPixel(row);

            // Set picture size to match the merged cell size (preserving aspect ratio)
            picture.Width = (int)totalWidth;
            picture.Height = (int)totalHeight;

            // Save the workbook
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
