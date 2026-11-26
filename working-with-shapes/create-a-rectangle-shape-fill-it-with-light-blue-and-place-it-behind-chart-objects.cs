// Title: Add a light‑blue rectangle shape behind a column chart in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to insert a rectangle shape behind a chart in Aspose.Cells C# | set solid fill color of a shape to light blue with Aspose.Cells .NET | change Z‑order of worksheet shapes to send a rectangle to back in Aspose.Cells | remove border line from a rectangle shape in Aspose.Cells C# example | create a column chart and add a background rectangle using Aspose.Cells
// Tags: rectangle shape background for chart Aspose.Cells | solid light blue fill Aspose.Cells shape | shape Z-order positioning Aspose.Cells | remove shape line border Aspose.Cells | column chart with background rectangle Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsRectangleBehindChart
{
    // The example creates a new workbook, adds sample data, generates a column chart, inserts a rectangle shape, fills it with a solid light‑blue color, removes its border, sets its Z‑order to place it behind the chart, and saves the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate sample data for the chart
                // -------------------------------------------------
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 6; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // -------------------------------------------------
                // 2. Add a column chart and set its data range
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.SetChartDataRange("A1:B6", true);
                chart.Title.Text = "Sample Column Chart";

                // -------------------------------------------------
                // 3. Add a rectangle shape behind the chart
                // -------------------------------------------------
                // Parameters: upper left row, top offset (pixels), upper left column, left offset (pixels), height (pixels), width (pixels)
                RectangleShape rect = sheet.Shapes.AddRectangle(2, 0, 2, 0, 150, 300);

                // Fill the rectangle with light blue color
                rect.Fill.FillType = FillType.Solid;
                rect.Fill.SolidFill.Color = Color.LightBlue;

                // Remove the line border
                rect.HasLine = false;

                // Set Z-order so the rectangle is behind other objects (charts have higher default Z-order)
                rect.ZOrderPosition = 0;

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                string outputPath = "RectangleBehindChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
