// Title: Copy Chart Shapes Between Excel Workbooks with Aspose.Cells for .NET
// Description: A C# utility that loads a source Excel file, creates a destination workbook, copies each worksheet's cells and formatting, then transfers only ChartShape objects while preserving their original size, style, and position, and finally saves the result.
// Keywords: Aspose.Cells | C# chart shape copy | copy Excel charts programmatically | preserve chart formatting Aspose | chart shape collection | AddCopy Aspose.Cells | Excel workbook chart transfer
// Common Searches: How to copy chart shapes from one Excel file to another using Aspose.Cells | Copy Excel charts while keeping their position in C# | Aspose.Cells copy only ChartShape objects | Transfer chart formatting between workbooks .NET | Duplicate charts across worksheets with Aspose.Cells
// Developer Intent: Copy all ChartShape objects from a source workbook to a destination workbook, retaining their formatting and placement.
// Use Cases: Create client‑specific reports by reusing chart layouts from a master template. | Consolidate charts from multiple source files into a single summary workbook. | Automate dashboard generation where chart positions must stay unchanged. | Migrate legacy Excel dashboards to a new workbook without manual redesign.
// AI Prompts: Write C# code that uses Aspose.Cells to copy ChartShape objects between workbooks while preserving size, style, and position. | Show how to modify the utility to update the data range of each copied chart to a new worksheet. | Provide a version that copies charts only from selected worksheet names passed as parameters. | Explain error handling for missing source files and how to log copied chart details.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// A C# utility that loads a source Excel file, creates a destination workbook, copies each worksheet's cells and formatting, then transfers only ChartShape objects while preserving their original size, style, and position, and finally saves the result.
class ChartShapeCopier
{
    // Copies all chart shapes from source workbook to destination workbook while preserving formatting.
    public static void CopyCharts(string sourcePath, string destinationPath)
    {
        // Verify source file exists.
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException($"Source file not found: {sourcePath}");

        // Load the source workbook.
        Workbook sourceWorkbook = new Workbook(sourcePath);

        // Create an empty destination workbook.
        Workbook destinationWorkbook = new Workbook();

        // Iterate through each worksheet in the source workbook.
        foreach (Worksheet sourceSheet in sourceWorkbook.Worksheets)
        {
            // Try to get a worksheet with the same name in the destination workbook.
            Worksheet destSheet = destinationWorkbook.Worksheets[sourceSheet.Name];

            // If it does not exist, add a new worksheet and set its name.
            if (destSheet == null)
            {
                int newIndex = destinationWorkbook.Worksheets.Add();
                destSheet = destinationWorkbook.Worksheets[newIndex];
                destSheet.Name = sourceSheet.Name;
            }

            // Copy cells, values, and formatting (charts are not copied by this method).
            destSheet.Copy(sourceSheet);

            // Get shape collections from both worksheets.
            ShapeCollection sourceShapes = sourceSheet.Shapes;
            ShapeCollection destShapes = destSheet.Shapes;

            // Iterate through source shapes and copy only chart shapes.
            foreach (Shape sourceShape in sourceShapes)
            {
                if (sourceShape is ChartShape)
                {
                    // Preserve the original position of the chart shape.
                    int topRow = sourceShape.UpperLeftRow;
                    int top = sourceShape.Top;
                    int leftColumn = sourceShape.UpperLeftColumn;
                    int left = sourceShape.Left;

                    // Add a copy of the chart shape to the destination worksheet.
                    destShapes.AddCopy(sourceShape, topRow, top, leftColumn, left);
                }
            }
        }

        // Ensure the directory for the destination file exists.
        string destDir = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
            Directory.CreateDirectory(destDir);

        // Save the destination workbook.
        destinationWorkbook.Save(destinationPath);
    }

    static void Main()
    {
        try
        {
            string sourceFile = "source.xlsx";
            string destinationFile = "destination.xlsx";

            CopyCharts(sourceFile, destinationFile);
            Console.WriteLine("All chart shapes have been copied successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
