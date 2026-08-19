// Title: Aspose.Cells .NET – Turn Off Chart Data Label Auto‑Fit and Apply Fixed Width
// Description: This C# example builds a workbook, inserts a column chart, enables data labels, then disables each label’s auto‑fit, assigns a constant pixel width, and locks the font size before saving the file.
// Keywords: Aspose.Cells | .NET | C# | chart data labels | disable auto‑fit | fixed label width | pixel width | ChartPoint.DataLabels | IsResizeShapeToFitText | AutoScaleFont false
// Common Searches: Aspose.Cells chart label fixed width | prevent data label resizing in Excel using Aspose | set constant pixel size for chart data labels C# | disable auto‑fit of data labels Aspose.Cells | keep font size unchanged in chart labels .NET
// Developer Intent: Stop chart data labels from automatically resizing and enforce a uniform width and font size.
// Use Cases: Design dashboards where label dimensions must stay consistent across varying values | Export Excel reports with predictable layout for printing or PDF conversion | Create templates that align labels with surrounding graphics or tables | Maintain visual uniformity in multi‑series charts for branding guidelines
// AI Prompts: Write C# code with Aspose.Cells that disables auto‑fit for chart data labels and sets a 60‑pixel width. | How can I keep chart data label shapes from resizing in an Aspose.Cells column chart? | Show the steps to fix data label width and prevent font scaling in Aspose.Cells for .NET. | Provide an example that iterates over ChartPoint objects to set DataLabels.IsResizeShapeToFitText = false.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelAutoFitDemo
{
    // This C# example builds a workbook, inserts a column chart, enables data labels, then disables each label’s auto‑fit, assigns a constant pixel width, and locks the font size before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart data source
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Disable auto‑fit for each data label and set a fixed width (in pixels)
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsResizeShapeToFitText = false; // Prevent shape from auto‑resizing
                point.DataLabels.WidthPixel = 60;               // Consistent label width
                point.DataLabels.AutoScaleFont = false;         // Keep font size unchanged
            }

            // Save the workbook
            workbook.Save("DataLabelAutoFitDisabled.xlsx");
        }
    }
}
