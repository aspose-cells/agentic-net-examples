// Title: Disable auto‑fit for chart data labels and set a fixed pixel width with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, enables data labels, then loops through each ChartPoint to turn off IsResizeShapeToFitText and assign a constant WidthPixel value, preventing automatic label resizing and keeping label widths uniform before saving the file.
// Keywords: Aspose.Cells data label auto fit | fixed width chart labels .NET | IsResizeShapeToFitText Aspose | WidthPixel chart data labels | disable chart label resizing | Aspose.Cells C# chart customization
// Common Searches: how to turn off auto‑fit for data labels in Aspose.Cells | set fixed pixel width for chart data labels C# | prevent data label shape from resizing Aspose.Cells chart | Aspose.Cells chart label size control
// Developer Intent: The developer wants to stop chart data labels from automatically resizing and apply a consistent pixel width to each label.
// Use Cases: Designing dashboards where all data labels share the same width for a clean layout. | Generating Excel reports that must follow corporate branding with uniform label dimensions. | Exporting charts to PDF or images while preserving consistent label sizes across rendering engines.
// AI Prompts: Show C# code to disable data label auto‑fit and set a fixed WidthPixel for each point in an Aspose.Cells chart. | Provide an example that iterates over chart points and sets IsResizeShapeToFitText = false and WidthPixel = 70 using Aspose.Cells. | Explain how to control data label size globally for a series in Aspose.Cells without looping through each point.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDataLabelAutoFitDemo
{
    // Creates a workbook, adds a column chart, enables data labels, then loops through each ChartPoint to turn off IsResizeShapeToFitText and assign a constant WidthPixel value, preventing automatic label resizing and keeping label widths uniform before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;
            series.DataLabels.Position = LabelPositionType.Center;

            // Disable auto‑fit (auto‑resize) for each data label to keep a consistent width
            foreach (ChartPoint point in series.Points)
            {
                // Prevent the label shape from resizing to fit its text
                point.DataLabels.IsResizeShapeToFitText = false;

                // Optionally set a fixed width (in pixels) for the label
                point.DataLabels.WidthPixel = 60;
            }

            // Save the workbook to an XLSX file
            workbook.Save("DataLabelAutoFitDisabled.xlsx");
        }
    }
}
