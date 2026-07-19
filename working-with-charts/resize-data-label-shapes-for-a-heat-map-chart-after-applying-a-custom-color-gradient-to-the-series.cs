// Title: Resize Data Label Shapes in a Surface3D Heat‑Map Chart with Gradient Fill – Aspose.Cells for .NET
// Description: Creates a workbook with a matrix, adds a Surface3D chart, applies a vertical two‑color gradient to the series, enables centered data labels, disables automatic shape resizing, sets each label to a fixed width of 50 pt and height of 20 pt, recalculates the chart, and saves the file.
// Keywords: Aspose.Cells resize data label | Surface3D chart gradient fill | C# fixed data label size | heat map chart Aspose.Cells | .NET chart label dimensions | two‑color vertical gradient Aspose
// Common Searches: how to set fixed width for chart data labels Aspose.Cells | apply vertical gradient to Surface3D series .NET | disable auto‑fit for data label shapes in Excel chart | custom size heat map data labels C# | Aspose.Cells resize data label after gradient fill
// Developer Intent: Set a gradient fill on a Surface3D chart series and give every data label a uniform, manually defined size.
// Use Cases: Produce a heat‑map matrix where all value labels share the same dimensions for a clean layout. | Highlight low vs. high values with a two‑tone gradient while keeping label sizes consistent. | Generate Excel reports that require non‑resizable data labels to align with surrounding graphics.
// AI Prompts: Generate C# code with Aspose.Cells that adds a vertical two‑color gradient to a Surface3D chart and forces each data label to 60 pt width and 25 pt height. | Show how to turn off automatic data label resizing and assign a fixed shape size in an Aspose.Cells heat‑map chart. | Explain the steps to recalculate a chart after modifying data label dimensions so the changes appear in the saved workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsHeatMapDataLabelResize
{
    // Creates a workbook with a matrix, adds a Surface3D chart, applies a vertical two‑color gradient to the series, enables centered data labels, disables automatic shape resizing, sets each label to a fixed width of 50 pt and height of 20 pt, recalculates the chart, and saves the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a heat‑map‑like matrix
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Q1");
                sheet.Cells["C1"].PutValue("Q2");
                sheet.Cells["D1"].PutValue("Q3");
                sheet.Cells["E1"].PutValue("Q4");

                string[] categories = { "North", "South", "East", "West" };
                double[,] values = {
                    { 10, 20, 30, 40 },
                    { 25, 15, 35, 45 },
                    { 20, 30, 10, 50 },
                    { 40, 20, 25, 35 }
                };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                    for (int j = 0; j < values.GetLength(1); j++)
                    {
                        sheet.Cells[i + 2, j + 1].PutValue(values[i, j]); // Columns B‑E
                    }
                }

                // Add a chart (Surface3D works for matrix data)
                int chartIndex = sheet.Charts.Add(ChartType.Surface3D, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:E5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Apply a two‑color vertical gradient to the series fill
                chart.NSeries[0].Area.FillFormat.SetTwoColorGradient(
                    Color.LightBlue,
                    Color.DarkBlue,
                    GradientStyleType.Vertical,
                    1);

                // Enable data labels and set their position
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true;
                series.DataLabels.Position = LabelPositionType.Center;

                // Resize each data label shape
                foreach (ChartPoint point in series.Points)
                {
                    point.DataLabels.IsResizeShapeToFitText = false;
                    point.DataLabels.Width = 50;   // custom width (points)
                    point.DataLabels.Height = 20;  // custom height (points)
                }

                // Recalculate the chart to apply changes
                chart.Calculate();

                // Save the workbook
                string outputPath = "HeatMapDataLabelResize.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
