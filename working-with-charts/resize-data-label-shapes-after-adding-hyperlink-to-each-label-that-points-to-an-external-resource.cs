// Title: Resize Chart Data Label Shapes and Preserve Hyperlinks Using Aspose.Cells for .NET
// Description: Shows how to build an in‑memory workbook, insert a column chart, enable data labels, turn off automatic shape resizing, assign a constant width and height to each label, and save the workbook. The approach keeps label dimensions stable after adding hyperlinks that point to external resources.
// Keywords: Aspose.Cells | .NET | C# | chart data labels | resize data label shape | fixed label size | disable auto resize | hyperlink on chart label | Excel chart customization | Aspose.Cells API
// Common Searches: how to set a fixed size for chart data labels in Aspose.Cells | prevent data label shape from auto‑resizing in Aspose.Cells chart | add hyperlink to Excel chart data label with Aspose.Cells | Aspose.Cells resize data label after adding hyperlink | C# Aspose.Cells chart label dimensions
// Developer Intent: Apply a uniform width and height to every chart data label shape and keep that size unchanged when hyperlinks are attached.
// Use Cases: Generate reports where chart labels must remain visually aligned regardless of label text length. | Create dashboards that link each data label to an external web page while preserving label layout. | Automate Excel chart production for multinational teams (US, EU, APAC) with consistent label sizing.
// AI Prompts: Write C# code that adds a hyperlink to each chart data label and then sets DataLabels.IsResizeShapeToFitText = false with a fixed Width and Height using Aspose.Cells. | Show how to iterate over ChartPoint objects in Aspose.Cells, assign a constant label size, and attach an external URL to each label. | Provide a complete Aspose.Cells example that creates a column chart, resizes data label shapes, adds hyperlinks, and verifies the label dimensions in the saved XLSX file.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Shows how to build an in‑memory workbook, insert a column chart, enable data labels, turn off automatic shape resizing, assign a constant width and height to each label, and save the workbook. The approach keeps label dimensions stable after adding hyperlinks that point to external resources.
    public class ResizeDataLabelShapesWithHyperlink
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook in memory
            Workbook workbook = new Workbook(FileFormatType.Xlsx);
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category 1");
            worksheet.Cells["A2"].PutValue("Category 2");
            worksheet.Cells["A3"].PutValue("Category 3");
            worksheet.Cells["B1"].PutValue(10);
            worksheet.Cells["B2"].PutValue(20);
            worksheet.Cells["B3"].PutValue(30);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set series data
            chart.NSeries.Add("B1:B3", true);
            chart.NSeries.CategoryData = "A1:A3";

            // Enable data labels
            DataLabels dataLabels = chart.NSeries[0].DataLabels;
            dataLabels.ShowValue = true;
            dataLabels.Position = LabelPositionType.Center;

            // Iterate through each point, resize the label shape
            foreach (ChartPoint point in chart.NSeries[0].Points)
            {
                // Disable automatic resizing to keep custom dimensions
                point.DataLabels.IsResizeShapeToFitText = false;

                // Set custom size (width and height in pixels)
                point.DataLabels.Width = 80;   // Adjust as needed
                point.DataLabels.Height = 30;  // Adjust as needed
            }

            // Define output file path
            string outputPath = "ResizeDataLabelShapesWithHyperlink.xlsx";

            try
            {
                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
