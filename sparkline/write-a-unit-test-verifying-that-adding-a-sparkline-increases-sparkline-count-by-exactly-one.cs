using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsTests
{
    public class SparklineDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the sparkline
                worksheet.Cells["A1"].PutValue(5);
                worksheet.Cells["B1"].PutValue(2);
                worksheet.Cells["C1"].PutValue(1);
                worksheet.Cells["D1"].PutValue(3);

                // Define the location where the sparkline will be placed
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 4,
                    EndColumn = 4
                };

                // Add a sparkline group (type Line) with the data range and location
                int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
                SparklineGroup group = worksheet.SparklineGroups[groupIndex];

                // Record the initial count (should be 0)
                int initialCount = group.Sparklines.Count;

                // Add a sparkline to the collection
                group.Sparklines.Add("A1:D1", 0, 4);

                // Record the count after adding
                int finalCount = group.Sparklines.Count;

                // Verify that the count increased by exactly one
                if (finalCount == initialCount + 1)
                {
                    Console.WriteLine("Sparkline count increased by one as expected.");
                }
                else
                {
                    Console.WriteLine($"Unexpected sparkline count. Initial: {initialCount}, Final: {finalCount}");
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "SparklineAddTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}