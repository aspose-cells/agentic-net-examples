// Title: Create a line sparkline in column B (cell B1) from range A1:A10 using Aspose.Cells for .NET (C#)
// Common Searches: c# aspso cells add a line sparkline to cell B1 using range A1:A10 | how to use SparklineGroup.Add to create a line sparkline in Excel with Aspose.Cells | example of SparklineCollection.Add for a single sparkline in column B in C# | setting high point low point first and last point for a sparkline with Aspose.Cells | saving an Excel workbook with a line sparkline as .xlsx using Aspose.Cells
// Tags: sparklinegroup.add c# | sparklinecollection.add data range A1:A10 | excel sparkline column B aspso cells | configure sparkline visual options aspso cells | save workbook xlsx with sparkline aspso cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineDemo
{
    // The sample creates a new workbook, fills cells A1:A10 with incremental values, adds a line SparklineGroup, inserts a sparkline into cell B1 using the same range, enables high/low/first/last point markers, and saves the result as SparklineLine_B1.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data in column A (A1:A10)
                for (int i = 0; i < 10; i++)
                {
                    // Cells are zero‑based indexed: row i, column 0 (A)
                    worksheet.Cells[i, 0].PutValue(i + 1); // simple incremental data
                }

                // Define the location where the sparkline will be placed: cell B1
                // CellArea uses zero‑based indexes as well
                CellArea sparklineLocation = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 1, // Column B
                    EndColumn = 1
                };

                // Add a sparkline group of type Line.
                // Parameters: SparklineType, data range for the group (can be a superset), 
                // isVertical (false for horizontal), and the location area.
                int groupIndex = worksheet.SparklineGroups.Add(
                    SparklineType.Line,
                    "A1:A10",          // group data range (can be same as individual sparkline)
                    false,             // not vertical
                    sparklineLocation  // where the sparklines will be placed
                );

                // Retrieve the created SparklineGroup
                SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

                // Add a single sparkline to the group.
                // Signature: Add(string dataRange, int row, int column)
                // Data range is the source data for this sparkline.
                // Row and column specify the cell that will contain the sparkline.
                sparklineGroup.Sparklines.Add("A1:A10", 0, 1); // B1

                // Optional: set some visual properties for demonstration
                sparklineGroup.ShowHighPoint = true;
                sparklineGroup.ShowLowPoint = true;
                sparklineGroup.ShowFirstPoint = true;
                sparklineGroup.ShowLastPoint = true;

                // Save the workbook to a file
                string outputPath = "SparklineLine_B1.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine($"Sparkline added to column B (cell B1) using data range A1:A10.");
            }
            catch (Exception ex)
            {
                // Report any errors that occur during processing
                Console.WriteLine("An error occurred while creating the sparkline:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
