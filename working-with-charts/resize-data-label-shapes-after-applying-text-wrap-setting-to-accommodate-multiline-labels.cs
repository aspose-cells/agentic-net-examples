// Title: Resize chart data label shapes for wrapped multi‑line text using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a column chart, enable text wrap on data labels, set a fixed label width, and auto‑resize the label shape so multi‑line text fits correctly in an Excel workbook generated with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart data labels | resize data label shape | text wrap chart labels .NET | auto‑fit data label height | fixed label width pixel | multi‑line chart labels | Aspose.Cells example | Excel chart label formatting
// Common Searches: how to wrap text in Aspose.Cells chart data labels | auto resize data label shape after wrapping text Aspose.Cells | set fixed width for chart data labels C# | Aspose.Cells multi‑line data label example | adjust chart label size programmatically .NET
// Developer Intent: Programmatically enable text wrapping and auto‑resize for chart data label shapes so they expand vertically to accommodate multi‑line content.
// Use Cases: Create a column chart with long category names and ensure labels display the full text without truncation. | Apply consistent wrap and auto‑resize settings to every point in a series for uniform label appearance. | Generate Excel reports where label width is constrained (e.g., 80 px) and height adapts automatically to wrapped text.
// AI Prompts: Show C# code to enable text wrap and auto‑resize for chart data labels in Aspose.Cells. | Provide an Aspose.Cells example that sets a fixed label width and forces multi‑line labels to fit. | Explain how to adjust the shape size of data labels after enabling text wrapping in a .NET chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a column chart, enable text wrap on data labels, set a fixed label width, and auto‑resize the label shape so multi‑line text fits correctly in an Excel workbook generated with Aspose.Cells for .NET.
    public class ResizeDataLabelShapesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Long Category Name 1");
                sheet.Cells["A3"].PutValue("Long Category Name 2");
                sheet.Cells["A4"].PutValue("Long Category Name 3");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["B4"].PutValue(65);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                DataLabels seriesLabels = chart.NSeries[0].DataLabels;
                seriesLabels.ShowValue = true;
                seriesLabels.IsTextWrapped = true;               // Enable text wrapping
                seriesLabels.IsResizeShapeToFitText = true;      // Auto‑fit shape to wrapped text
                seriesLabels.WidthPixel = 80;                    // Fixed width to see wrapping effect

                // Ensure each point's individual label also respects the settings
                foreach (ChartPoint pt in chart.NSeries[0].Points)
                {
                    pt.DataLabels.IsTextWrapped = true;
                    pt.DataLabels.IsResizeShapeToFitText = true;
                    pt.DataLabels.WidthPixel = 80;   // same width as series label
                }

                // Save the workbook
                string outputPath = "ResizeDataLabelShapesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
