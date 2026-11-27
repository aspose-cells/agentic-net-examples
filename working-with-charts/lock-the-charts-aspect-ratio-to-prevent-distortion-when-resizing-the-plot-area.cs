// Title: Lock a chart’s aspect ratio and manually resize its plot area using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# lock chart aspect ratio when resizing | set IsLockAspectRatio property on chart shape in Aspose.Cells | prevent distortion of Excel chart plot area with Aspose.Cells .NET | manual plot area width and height with fixed aspect ratio Aspose.Cells | resize Aspose.Cells chart without changing its aspect ratio
// Tags: Aspose.Cells chart aspect ratio lock | C# chart shape IsLockAspectRatio | manual plot area sizing Aspose.Cells | prevent chart distortion Excel .NET | chart object resize preserving ratio

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartAspectRatioLock
{
    // The example creates a workbook, adds a column chart, accesses the chart's underlying Shape via ChartObject, sets IsLockAspectRatio to true, disables automatic plot area sizing, assigns explicit width and height that maintain a 4:3 ratio, and demonstrates that resizing the chart shape preserves the aspect ratio before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output Excel file
            const string outputPath = "ChartAspectRatioLocked.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Populate sample data for the chart
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B4"].PutValue(30);

            // ------------------------------------------------------------
            // 2. Add a column chart to the worksheet
            // ------------------------------------------------------------
            // Parameters: chart type, upper‑left row, upper‑left column,
            // lower‑right row, lower‑right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories
            chart.Title.Text = "Sample Column Chart";

            // ------------------------------------------------------------
            // 3. Lock the chart's aspect ratio
            // ------------------------------------------------------------
            // A chart in Aspose.Cells is represented internally as a shape.
            // The Shape class provides the IsLockAspectRatio property.
            // The Chart object exposes its underlying shape via the ChartObject property.
            // Setting IsLockAspectRatio to true prevents distortion when the chart
            // (or its plot area) is resized.
            Shape chartShape = chart.ChartObject;
            chartShape.IsLockAspectRatio = true;

            // ------------------------------------------------------------
            // 4. Resize the plot area while preserving the aspect ratio
            // ------------------------------------------------------------
            // Disable automatic sizing so that manual dimensions are respected.
            chart.PlotArea.IsAutomaticSize = false;

            // Set an initial width and height that respect the desired aspect ratio.
            // For example, a 4:3 ratio (width : height = 4 : 3).
            const double aspectRatio = 4.0 / 3.0;
            const int plotAreaHeight = 2000; // in points (1/4000 of chart area)
            int plotAreaWidth = (int)(plotAreaHeight * aspectRatio);

            chart.PlotArea.Width = plotAreaWidth;
            chart.PlotArea.Height = plotAreaHeight;

            // Because the shape's aspect ratio is locked, any subsequent resize
            // of the chart (e.g., moving the chart object) will keep the 4:3 ratio.
            // To demonstrate, we also resize the whole chart object.
            // The shape will automatically adjust its width/height to keep the ratio.
            chartShape.Width = 5000; // arbitrary value; height will be calculated
            // Height is automatically adjusted because IsLockAspectRatio = true

            // ------------------------------------------------------------
            // 5. Save the workbook
            // ------------------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
            Console.WriteLine("The chart's aspect ratio is locked; resizing the chart will not distort it.");
        }
    }
}
