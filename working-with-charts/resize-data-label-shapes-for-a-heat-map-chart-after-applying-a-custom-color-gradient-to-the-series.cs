// Title: Resize data label shapes in a heat‑map chart after applying a custom gradient using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set fixed width and height for chart data labels | how to apply a two‑color gradient to a series and keep data label size | disable automatic resizing of data label shapes in a heat map chart Aspose.Cells | center bold data labels on a column chart with custom gradient Aspose.Cells
// Tags: two‑color gradient fill series Aspose.Cells | fixed data label width height Aspose.Cells | disable data label auto‑resize Aspose.Cells | heat map chart data label formatting C# | column chart gradient series Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace HeatMapDataLabelResize
{
    // The example creates a workbook, populates a 4×4 matrix, adds a column chart used as a heat‑map, applies a vertical two‑color gradient to the series, enables centered data labels, disables their automatic resizing, sets each label to a fixed 40 × 20 point size with bold black text, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Populate sample data for the heat map (a 4x4 matrix)
                // ------------------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Q1");
                sheet.Cells["C1"].PutValue("Q2");
                sheet.Cells["D1"].PutValue("Q3");
                sheet.Cells["E1"].PutValue("Q4");

                // Category column
                sheet.Cells["A2"].PutValue("Product A");
                sheet.Cells["A3"].PutValue("Product B");
                sheet.Cells["A4"].PutValue("Product C");
                sheet.Cells["A5"].PutValue("Product D");

                // Random numeric values
                Random rnd = new Random();
                for (int row = 2; row <= 5; row++)
                {
                    for (int col = 2; col <= 5; col++)
                    {
                        sheet.Cells[row, col].PutValue(rnd.Next(10, 100));
                    }
                }

                // ------------------------------------------------------------
                // 2. Add a chart (using Column chart as a substitute for HeatMap)
                // ------------------------------------------------------------
                // Parameters: type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (numeric matrix B2:E5)
                chart.NSeries.Add("B2:E5", true);
                // Category data (optional for heat map, but we set it for completeness)
                chart.NSeries.CategoryData = "A2:A5";

                // ------------------------------------------------------------
                // 3. Apply a custom linear color gradient to the series
                // ------------------------------------------------------------
                Series series = chart.NSeries[0];
                // Use SetTwoColorGradient to define a linear gradient fill
                // GradientStyleType.Vertical with a variant of 1 creates a vertical gradient
                series.Area.FillFormat.SetTwoColorGradient(Color.LightYellow, Color.OrangeRed, GradientStyleType.Vertical, 1);

                // ------------------------------------------------------------
                // 4. Enable data labels and resize their shapes
                // ------------------------------------------------------------
                // Show the numeric value inside each cell of the chart
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Iterate over each point (cell) in the series and adjust the label shape
                foreach (ChartPoint point in series.Points)
                {
                    // Prevent the label shape from auto‑resizing to fit the text
                    point.DataLabels.IsResizeShapeToFitText = false;

                    // Set a fixed size that works well with the chart cells
                    point.DataLabels.Width = 40;   // width in points
                    point.DataLabels.Height = 20;  // height in points

                    // Optional: center the text and make it bold for readability
                    point.DataLabels.Font.IsBold = true;
                    point.DataLabels.Font.Color = Color.Black;
                    point.DataLabels.TextHorizontalAlignment = TextAlignmentType.Center;
                    point.DataLabels.TextVerticalAlignment = TextAlignmentType.Center;
                }

                // ------------------------------------------------------------
                // 5. Save the workbook
                // ------------------------------------------------------------
                string outFile = "HeatMap_With_Resized_DataLabels.xlsx";
                workbook.Save(outFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
