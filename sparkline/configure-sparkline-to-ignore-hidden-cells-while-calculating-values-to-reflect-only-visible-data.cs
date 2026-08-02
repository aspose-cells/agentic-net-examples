using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineDemo
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

                // Populate sample data (row 1 visible, row 2 hidden)
                sheet.Cells["A1"].PutValue(5);
                sheet.Cells["B1"].PutValue(2);
                sheet.Cells["C1"].PutValue(1);
                sheet.Cells["D1"].PutValue(3);

                sheet.Cells["A2"].PutValue(10);
                sheet.Cells["B2"].PutValue(8);
                sheet.Cells["C2"].PutValue(6);
                sheet.Cells["D2"].PutValue(4);

                // Hide the second row – its values should be ignored by the sparkline
                sheet.Cells.Rows[1].IsHidden = true; // zero‑based index

                // Define the location range for the sparklines (E1:E2)
                CellArea sparklineLocation = new CellArea
                {
                    StartColumn = 4, // Column E (0‑based)
                    EndColumn = 4,
                    StartRow = 0,    // Row 1
                    EndRow = 1       // Row 2 (to match two data rows)
                };

                // Add a sparkline group that uses the data range A1:D2 (both rows)
                int groupIndex = sheet.SparklineGroups.Add(
                    SparklineType.Line,
                    "A1:D2",
                    false,
                    sparklineLocation);

                SparklineGroup group = sheet.SparklineGroups[groupIndex];

                // Exclude hidden cells from calculations
                group.DisplayHidden = false;

                // Optional: set a series color for better visibility
                CellsColor seriesColor = workbook.CreateCellsColor();
                seriesColor.Color = Color.Blue;
                group.SeriesColor = seriesColor;

                // Save the workbook
                string outputPath = "SparklineIgnoreHidden.xlsx";
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