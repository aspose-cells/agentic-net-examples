// Title: Generate a Euclidean distance matrix for all shapes in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to retrieve absolute row and column of a shape in Aspose.Cells C# | calculate distances between shapes in an Excel file with Aspose.Cells | create a distance matrix worksheet for shapes using Aspose.Cells .NET | write shape distance calculations to a new sheet in Aspose.Cells | iterate all shapes across worksheets and get their positions Aspose.Cells | save workbook after adding a distance matrix sheet Aspose.Cells C#
// Tags: extract shape absolute coordinates Aspose.Cells | compute Euclidean distance matrix for Excel shapes | populate distance matrix worksheet Aspose.Cells | add analysis sheet with shape distances Aspose.Cells | iterate shapes across multiple worksheets Aspose.Cells | save workbook with new sheet Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeDistanceMatrix
{
    // The example loads an existing workbook, gathers each shape's name and absolute row/column positions (UpperLeftRow, UpperLeftColumn) from all worksheets, creates a new sheet named "DistanceMatrix", writes shape names as headers, computes Euclidean distances between every pair of shapes, fills the matrix with these values, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string sourceFile = "InputWorkbook.xlsx";
            string outputFile = "OutputWorkbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule: use load)
                Workbook workbook = new Workbook(sourceFile);

                // Collect all shapes from all worksheets
                List<ShapeInfo> shapes = new List<ShapeInfo>();
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Retrieve absolute position (row, column). Offsets are omitted for compatibility.
                        double rowPos = shape.UpperLeftRow;
                        double colPos = shape.UpperLeftColumn;

                        // Store shape information
                        shapes.Add(new ShapeInfo
                        {
                            Name = shape.Name ?? string.Empty,
                            SheetName = sheet.Name,
                            Row = rowPos,
                            Column = colPos
                        });
                    }
                }

                // Prepare a new worksheet for the distance matrix
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet matrixSheet = workbook.Worksheets[newSheetIndex];
                matrixSheet.Name = "DistanceMatrix";

                // Write header (shape names) in first row and first column
                for (int i = 0; i < shapes.Count; i++)
                {
                    // Column headers (first row, starting from column B)
                    matrixSheet.Cells[0, i + 1].PutValue(shapes[i].Name);
                    // Row headers (first column, starting from row 2)
                    matrixSheet.Cells[i + 1, 0].PutValue(shapes[i].Name);
                }

                // Calculate Euclidean distances and fill the matrix
                for (int i = 0; i < shapes.Count; i++)
                {
                    for (int j = 0; j < shapes.Count; j++)
                    {
                        double distance = i == j ? 0.0 :
                            Math.Sqrt(
                                Math.Pow(shapes[i].Row - shapes[j].Row, 2) +
                                Math.Pow(shapes[i].Column - shapes[j].Column, 2));

                        // Write distance to cell (row i+2, column j+2) because of headers
                        matrixSheet.Cells[i + 1, j + 1].PutValue(distance);
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook (lifecycle rule: use save)
                workbook.Save(outputFile);
                Console.WriteLine($"Distance matrix saved to {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Helper class to store shape position information
        private class ShapeInfo
        {
            public string Name { get; set; } = string.Empty;
            public string SheetName { get; set; } = string.Empty;
            public double Row { get; set; }      // Absolute row position
            public double Column { get; set; }   // Absolute column position
        }
    }
}
