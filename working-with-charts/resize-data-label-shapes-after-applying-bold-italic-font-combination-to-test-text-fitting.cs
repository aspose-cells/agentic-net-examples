// Title: Auto‑Resize Chart Data Label Shapes for Bold & Italic Text with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a column chart in an Excel workbook, enable data labels, apply bold and italic formatting, and automatically resize each label shape to fit the styled text using the IsResizeShapeToFitText property in Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# chart data labels | auto resize data label shape | IsResizeShapeToFitText Aspose | bold italic font chart labels | Excel chart label auto‑fit .NET | ChartPoint DataLabels resize | column chart label sizing | Aspose.Cells example | global Excel automation
// Common Searches: Aspose.Cells auto‑resize chart data labels | C# set IsResizeShapeToFitText for chart points | fit bold italic text in Excel chart labels | how to enlarge data label shape Aspose.Cells | chart label auto‑fit after font change
// Developer Intent: Programmatically adjust chart data label shapes so they expand automatically to accommodate bold and italic font styling.
// Use Cases: Generating Excel reports where emphasized data labels must remain fully visible without manual resizing. | Creating dashboards that apply bold/italic styling to highlight key values and need labels to auto‑adjust. | Automating bulk spreadsheet creation where label dimensions vary based on dynamic content.
// AI Prompts: Show C# code to enable IsResizeShapeToFitText for all points in an Aspose.Cells chart after applying bold and italic fonts. | Explain the interaction between series.DataLabels.ApplyFont and point.DataLabels.IsResizeShapeToFitText in Aspose.Cells. | Provide a step‑by‑step guide to auto‑fit chart data label shapes to styled text using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a column chart in an Excel workbook, enable data labels, apply bold and italic formatting, and automatically resize each label shape to fit the styled text using the IsResizeShapeToFitText property in Aspose.Cells for .NET.
    public class ResizeDataLabelShapesDemo
    {
        public static void Run()
        {
            try
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
                series.DataLabels.Position = LabelPositionType.Center;

                // Apply bold and italic font to the data labels
                series.DataLabels.Font.IsBold = true;
                series.DataLabels.Font.IsItalic = true;
                series.DataLabels.Font.Size = 12;
                series.DataLabels.Font.Color = Color.DarkBlue;

                // Apply the font settings to all individual data label objects
                series.DataLabels.ApplyFont();

                // For each data point, enable auto‑resize of the label shape to fit the new font
                foreach (ChartPoint point in series.Points)
                {
                    // Allow the shape to auto‑fit the text
                    point.DataLabels.IsResizeShapeToFitText = true;

                    // Optionally set an initial small width to demonstrate resizing
                    point.DataLabels.Width = 40;
                }

                // Save the workbook
                workbook.Save("ResizeDataLabelShapesDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ResizeDataLabelShapesDemo.Run();
        }
    }
}
