// Title: Resize chart data label shapes to fit bold‑italic text in an Aspose.Cells column chart (C#)
// Common Searches: Aspose.Cells C# resize data label shape to fit long bold italic text in column chart | How to enable IsResizeShapeToFitText for chart data labels with custom text using Aspose.Cells .NET | C# Aspose.Cells set bold and italic font for chart data labels and auto‑fit shape size | Auto‑fit chart data label boxes after applying bold‑italic formatting in Aspose.Cells | Create column chart with long custom data labels that automatically resize in Aspose.Cells for .NET
// Tags: Aspose.Cells chart data label auto resize | C# column chart data label formatting | IsResizeShapeToFitText property usage | bold italic font for chart labels | custom long data label text handling

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example builds a column chart, applies a bold‑italic font to all data labels, enables IsResizeShapeToFitText, assigns lengthy custom text to each point, and saves the workbook as an XLSX file, demonstrating automatic resizing of data label shapes to accommodate formatted text.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // 1. Populate sample data for the chart
        // -------------------------------------------------
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Alpha");
        sheet.Cells["A3"].PutValue("Beta");
        sheet.Cells["A4"].PutValue("Gamma");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(15);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(45);

        // -------------------------------------------------
        // 2. Add a column chart
        // -------------------------------------------------
        // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data source for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // -------------------------------------------------
        // 3. Configure data labels for each point
        // -------------------------------------------------
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;               // Show the numeric value
        series.DataLabels.IsAutoText = false;             // We'll supply custom text
        series.DataLabels.IsTextWrapped = true;           // Allow wrapping if needed

        // Apply a bold‑italic font to all data labels
        series.DataLabels.Font.IsBold = true;
        series.DataLabels.Font.IsItalic = true;
        series.DataLabels.Font.Size = 12;
        series.DataLabels.Font.Color = Color.DarkBlue;

        // Ensure the shape of each data label resizes to fit the text
        series.DataLabels.IsResizeShapeToFitText = true;

        // Provide a long custom text for each point to test the auto‑fit behaviour
        for (int i = 0; i < series.Points.Count; i++)
        {
            ChartPoint point = series.Points[i];
            // Disable the default auto‑text for the point and set our own text
            point.DataLabels.IsAutoText = false;
            point.DataLabels.Text = $"Point {i + 1}: This is a very long label to test auto‑fit.";
            // Ensure the point inherits the same resize‑to‑fit setting
            point.DataLabels.IsResizeShapeToFitText = true;
        }

        // -------------------------------------------------
        // 4. Save the workbook
        // -------------------------------------------------
        string outFile = "ResizeDataLabelsDemo.xlsx";
        workbook.Save(outFile, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to {outFile}");
    }
}
