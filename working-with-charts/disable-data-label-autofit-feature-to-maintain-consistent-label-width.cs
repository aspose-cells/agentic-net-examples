// Title: How to turn off data label auto‑fit and set a fixed width for chart labels in Aspose.Cells for .NET
// Common Searches: Aspose.Cells disable data label auto fit for column chart C# | set fixed width for chart data labels using Aspose.Cells .NET | prevent automatic resizing of data labels in Aspose.Cells chart | chart point data label size customization Aspose.Cells C# example | how to use IsResizeShapeToFitText property in Aspose.Cells
// Tags: data label auto‑fit control Aspose.Cells | chart data label width customization .NET | column chart label size management C# | IsResizeShapeToFitText false usage Aspose.Cells | fixed label shape dimensions Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds a column chart with data labels, disables the auto‑fit feature for each label, assigns a fixed width to the label shapes, and saves the workbook as an .xlsx file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels and set basic properties
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;
        dataLabels.Position = LabelPositionType.Center;

        // Disable auto‑fit (auto‑resize) for each data label to keep a consistent width
        foreach (ChartPoint point in chart.NSeries[0].Points)
        {
            // Turn off the auto‑fit feature
            point.DataLabels.IsResizeShapeToFitText = false;

            // Optionally set a fixed width (in points) for the label shape
            point.DataLabels.Width = 60; // Adjust as needed
        }

        // Save the workbook to a file
        string outputPath = "ChartDataLabelAutoFitDisabled.xlsx";
        workbook.Save(outputPath);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
