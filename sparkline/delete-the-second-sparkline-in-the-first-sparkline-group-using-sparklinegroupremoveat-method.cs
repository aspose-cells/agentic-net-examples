using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineRemovalDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (two rows) that will be used for the sparklines
                // Row 1
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(8);
                sheet.Cells["D1"].PutValue(3);
                // Row 2
                sheet.Cells["A2"].PutValue(6);
                sheet.Cells["B2"].PutValue(1);
                sheet.Cells["C2"].PutValue(7);
                sheet.Cells["D2"].PutValue(4);

                // Define the location range for two sparklines (cells E1 and F1)
                // StartColumn = 4 (E), EndColumn = 5 (F), rows 0‑1 (E1:F2)
                CellArea location = new CellArea
                {
                    StartRow = 0,
                    EndRow = 1,
                    StartColumn = 4,
                    EndColumn = 5
                };

                // Add a sparkline group with the data range A1:D2.
                // The location range contains two cells, so two sparklines are created automatically.
                int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D2", false, location);
                SparklineGroup sparklineGroup = sheet.SparklineGroups[groupIndex];

                // At this point sparklineGroup.Sparklines.Count == 2
                Console.WriteLine($"Initial sparkline count: {sparklineGroup.Sparklines.Count}");

                // Delete the second sparkline (index 1) using RemoveAt on the Sparklines collection
                sparklineGroup.Sparklines.RemoveAt(1);

                // Verify removal
                Console.WriteLine($"Sparkline count after removal: {sparklineGroup.Sparklines.Count}");

                // Save the workbook
                string outputPath = "SparklineRemovalResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}