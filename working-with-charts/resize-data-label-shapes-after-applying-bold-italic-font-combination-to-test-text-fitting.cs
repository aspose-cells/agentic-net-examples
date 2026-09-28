// Title: How to auto‑resize chart data label shapes after applying bold and italic font in Aspose.Cells for .NET
// AI Prompts: Create a column chart, set the data label font to bold, italic, blue, size 12, and enable each label shape to auto‑resize to fit the text using Aspose.Cells in C#. | Update an existing Aspose.Cells chart so that after changing the data label font, the IsResizeShapeToFitText property is set to true for all points, ensuring label shapes adjust automatically. | Generate a workbook that demonstrates applying a bold‑italic font to a series' data labels and programmatically calling ApplyFont and IsResizeShapeToFitText to fit the label shapes.
// Common Searches: Aspose.Cells C# resize data label shape to fit bold italic text | how to enable auto‑fit for chart data labels after font change in Aspose.Cells | IsResizeShapeToFitText property example for column chart .NET | apply bold and italic font to all data labels and auto‑resize shapes Aspose.Cells
// Tags: auto‑resize chart data label shapes Aspose.Cells | apply bold italic font to chart data labels .NET | ChartPoint.DataLabels.IsResizeShapeToFitText usage | column chart data label formatting Aspose.Cells | ApplyFont method for series data labels

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates creating a column chart, applying a bold‑italic blue font to data labels, and enabling each data label shape to auto‑resize to fit the text using Aspose.Cells for .NET.
    public class ResizeDataLabelShapesDemo
    {
        public static void Main(string[] args)
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply bold and italic font to the data labels
            series.DataLabels.Font.IsBold = true;
            series.DataLabels.Font.IsItalic = true;
            series.DataLabels.Font.Size = 12;
            series.DataLabels.Font.Color = Color.Blue;

            // Propagate the font settings to all individual data label objects
            series.DataLabels.ApplyFont();

            // Ensure each data label shape auto‑fits the text after the font change
            foreach (ChartPoint point in series.Points)
            {
                // Enable auto‑resize to fit the text (default is true, set explicitly for clarity)
                point.DataLabels.IsResizeShapeToFitText = true;
            }

            // Save the workbook
            string outputPath = "ResizeDataLabelShapesDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
