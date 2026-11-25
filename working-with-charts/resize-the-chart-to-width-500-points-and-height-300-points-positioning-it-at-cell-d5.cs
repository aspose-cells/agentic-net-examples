// Title: How to resize an Aspose.Cells chart to 500 pt × 300 pt and position it at cell D5 using C#
// Common Searches: Aspose.Cells C# set chart width to 500 points and height to 300 points | C# place Aspose.Cells chart top‑left corner on cell D5 | Calculate X and Y coordinates in points for a chart based on column and row offsets in Aspose.Cells | Use PlacementType.MoveAndSize to make an Aspose.Cells chart move and resize with cells | Convert column/row pixel widths to points for chart placement in Aspose.Cells
// Tags: chart resizing points Aspose.Cells | chart positioning cell D5 Aspose.Cells | column offset calculation points Aspose.Cells | row offset calculation points Aspose.Cells | PlacementType.MoveAndSize Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartResizeAndPosition
{
    // The example creates a workbook, adds sample data, inserts a column chart, resizes it to 500 pt × 300 pt, computes the X/Y offsets in points for column D and row 5, positions the chart at cell D5, sets the chart to move and size with cells, and saves the file as ResizedChart.xlsx.
    class Program
    {
        private static double GetColumnOffsetInPoints(Worksheet worksheet, int columnIndex)
        {
            double offsetPixels = 0.0;
            for (int i = 0; i < columnIndex; i++)
            {
                // Width in pixels for column i
                offsetPixels += worksheet.Cells.GetColumnWidthPixel(i);
            }

            // Convert pixels to points (1 point = 1/72 inch, 1 pixel = 1/96 inch)
            return offsetPixels * 72.0 / 96.0;
        }

        /// <summary>
        /// Calculates the vertical offset (in points) of the top edge of the specified row.
        /// </summary>
        private static double GetRowOffsetInPoints(Worksheet worksheet, int rowIndex)
        {
            double offsetPixels = 0.0;
            for (int i = 0; i < rowIndex; i++)
            {
                // Height in pixels for row i
                offsetPixels += worksheet.Cells.GetRowHeightPixel(i);
            }

            // Convert pixels to points
            return offsetPixels * 72.0 / 96.0;
        }

        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and add some sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // -----------------------------------------------------------------
                // 2. Add a chart to the worksheet.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -----------------------------------------------------------------
                // 3. Resize the chart to 500 points (width) x 300 points (height).
                // -----------------------------------------------------------------
                chart.ChartObject.Width = 500;   // Width in points
                chart.ChartObject.Height = 300;  // Height in points

                // -----------------------------------------------------------------
                // 4. Position the chart so that its top‑left corner aligns with cell D5.
                // -----------------------------------------------------------------
                int targetColumn = 3; // Column D (0‑based)
                int targetRow = 4;    // Row 5 (0‑based)

                double offsetX = GetColumnOffsetInPoints(sheet, targetColumn);
                double offsetY = GetRowOffsetInPoints(sheet, targetRow);

                // ChartObject.X/Y expect integer values (points), so cast safely
                chart.ChartObject.X = (int)Math.Round(offsetX);
                chart.ChartObject.Y = (int)Math.Round(offsetY);

                // Make the chart move and size with cells
                chart.ChartObject.Placement = PlacementType.MoveAndSize;

                // -----------------------------------------------------------------
                // 5. Save the workbook.
                // -----------------------------------------------------------------
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ResizedChart.xlsx");
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to: {outputPath}");
                Console.WriteLine("Chart resized to 500pt x 300pt and positioned at cell D5.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
