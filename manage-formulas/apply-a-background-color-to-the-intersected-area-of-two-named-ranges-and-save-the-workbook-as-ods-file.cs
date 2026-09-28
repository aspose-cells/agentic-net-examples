// Title: Apply a yellow background to the intersecting cells of two CellArea ranges and export the workbook as ODS using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates two CellArea objects, calculates their overlapping region, applies a solid yellow fill style to each cell in that region, and saves the workbook in ODS format with Aspose.Cells. | Show how to programmatically determine the intersection of two named ranges, style the intersected cells with a background color, and write the result to an ODS file in a .NET application.
// Common Searches: how to highlight overlapping cells of two ranges using Aspose.Cells C# | Aspose.Cells compute intersection of CellArea and apply background color | save styled workbook as ODS with Aspose.Cells .NET | C# example for applying solid fill to intersected range in Excel file
// Tags: Aspose.Cells intersecting CellArea background fill | C# apply solid yellow style to range intersection | export workbook to ODS with cell styling Aspose.Cells | calculate CellArea overlap Aspose.Cells .NET | style intersected cells ODS export

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// // Creates a workbook, fills sample data, defines two CellArea ranges (B2:D5 and C3:E6), computes their intersection, applies a solid yellow background to the intersected cells, and saves the file as an ODS document.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data (optional, just to visualize the ranges)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define two cell areas directly
            CellArea area1 = CellArea.CreateCellArea("B2", "D5"); // B2:D5
            CellArea area2 = CellArea.CreateCellArea("C3", "E6"); // C3:E6

            // Compute the intersected area manually (avoids overload conflicts)
            CellArea intersectArea = new CellArea
            {
                StartRow = Math.Max(area1.StartRow, area2.StartRow),
                StartColumn = Math.Max(area1.StartColumn, area2.StartColumn),
                EndRow = Math.Min(area1.EndRow, area2.EndRow),
                EndColumn = Math.Min(area1.EndColumn, area2.EndColumn)
            };

            // Verify that there is an intersection
            if (intersectArea.StartRow <= intersectArea.EndRow &&
                intersectArea.StartColumn <= intersectArea.EndColumn)
            {
                // Create a style with yellow background
                Style intersectStyle = workbook.CreateStyle();
                intersectStyle.ForegroundColor = Color.Yellow;
                intersectStyle.Pattern = BackgroundType.Solid;

                // Apply the style to each cell in the intersected area
                for (int row = intersectArea.StartRow; row <= intersectArea.EndRow; row++)
                {
                    for (int col = intersectArea.StartColumn; col <= intersectArea.EndColumn; col++)
                    {
                        cells[row, col].SetStyle(intersectStyle);
                    }
                }
            }

            // Save the workbook as ODS
            string outputPath = "IntersectedRanges.ods";
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
