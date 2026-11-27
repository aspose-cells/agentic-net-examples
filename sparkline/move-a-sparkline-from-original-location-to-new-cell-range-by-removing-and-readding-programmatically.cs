// Title: How to relocate a sparkline to a different cell range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move sparkline to another cell range | programmatically change sparkline location in Excel with Aspose.Cells | remove sparkline and add it at new position using Aspose.Cells .NET | C# code to reposition a line sparkline in a workbook | Aspose.Cells example for moving sparkline between cells | update sparkline cell address in Excel via Aspose.Cells API
// Tags: sparkline relocation using Aspose.Cells API | reassign sparkline to new cell range C# | update sparkline data range cell address Aspose.Cells .NET | sparkline group cell address manipulation | move sparkline within worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineMover
{
    // The sample loads an existing workbook (or creates one with a line sparkline), finds the first sparkline in the first SparklineGroup, removes it, adds it back at a new row‑column location while keeping the original data range, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the input and output workbooks
            string inputPath = "SparklineDemo.xlsx";
            string outputPath = "SparklineMoved.xlsx";

            // Ensure the input file exists; if not, create a sample workbook with a sparkline
            if (!File.Exists(inputPath))
            {
                CreateSampleWorkbookWithSparkline(inputPath);
                Console.WriteLine($"Sample workbook created at '{inputPath}'.");
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Verify that the worksheet contains at least one sparkline group
                if (sheet.SparklineGroups.Count == 0)
                {
                    Console.WriteLine("No sparkline groups found in the worksheet. Exiting.");
                    return;
                }

                // For this example we work with the first sparkline group
                SparklineGroup group = sheet.SparklineGroups[0];

                // Verify that the group contains at least one sparkline
                if (group.Sparklines.Count == 0)
                {
                    Console.WriteLine("The sparkline group does not contain any sparklines. Exiting.");
                    return;
                }

                // Get the first sparkline to move
                Sparkline sparklineToMove = group.Sparklines[0];

                // Retrieve the data range of the sparkline (e.g., "A1:D1")
                string dataRange = sparklineToMove.DataRange;

                // Store the original location (row, column) for logging
                int originalRow = sparklineToMove.Row;
                int originalColumn = sparklineToMove.Column;

                // Remove the sparkline from its current location
                group.Sparklines.RemoveSparkline(sparklineToMove);
                Console.WriteLine($"Removed sparkline from Row={originalRow}, Column={originalColumn}.");

                // Define the new location for the sparkline (e.g., row 5, column 6 -> cell F6)
                int newRow = 5;      // zero‑based index (row 6 in Excel)
                int newColumn = 5;   // zero‑based index (column F in Excel)

                // Add the sparkline back with the same data range but at the new location
                int newSparklineIndex = group.Sparklines.Add(dataRange, newRow, newColumn);
                Console.WriteLine($"Added sparkline at new location Row={newRow}, Column={newColumn} (index {newSparklineIndex}).");

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved with moved sparkline to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while moving the sparkline:");
                Console.WriteLine(ex.Message);
            }
        }

        /// <param name="filePath">Full path where the workbook will be saved.</param>
        private static void CreateSampleWorkbookWithSparkline(string filePath)
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data in A1:D1
            ws.Cells["A1"].PutValue(5);
            ws.Cells["B1"].PutValue(2);
            ws.Cells["C1"].PutValue(8);
            ws.Cells["D1"].PutValue(3);

            // Define the cell area where the sparkline will be placed (E1)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group of type Line with the data range A1:D1
            int groupIndex = ws.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
            SparklineGroup group = ws.SparklineGroups[groupIndex];

            // Add the sparkline to the group (the Add method also creates the sparkline at the location)
            group.Sparklines.Add("A1:D1", 0, 4);

            // Save the workbook
            wb.Save(filePath, SaveFormat.Xlsx);
        }
    }
}
