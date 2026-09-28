// Title: Create a pie chart with percentage data labels sourced from worksheet cells and auto‑resize label shapes using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds an Excel workbook, fills cells A2:A4 and B2:B4 with categories and values, adds a pie chart linked to those ranges, enables data labels to show only percentages, and configures the label boxes to auto‑fit the text before saving the file. | Generate a complete Aspose.Cells example that inserts a pie chart, displays percentage labels derived from the source cells, and adjusts the label shape size dynamically to accommodate the label content.
// Common Searches: asp.net how to display only percentage labels on a pie chart with Aspose.Cells | c# Aspose.Cells auto resize data label shape to fit text in pie chart | example binding pie chart series to worksheet cell range using Aspose.Cells | set data label showpercentage property Aspose.Cells C# tutorial | adjust pie chart data label dimensions programmatically Aspose.Cells
// Tags: pie chart percentage data labels Aspose.Cells | auto‑fit data label shape Aspose.Cells | bind chart series to cell range C# | set data label dimensions Aspose.Cells | generate pie chart from worksheet data Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // This example creates a new workbook, populates it with category and value data, adds a pie chart linked to those cells, enables data labels to show only percentages, and sets the label shapes to automatically resize to fit the label text before saving as PieChartWithPercentLabels.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate worksheet with sample data for the pie chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a pie chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Pie, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series and show percentages
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowPercentage = true;   // display percentage values
        dataLabels.ShowValue = false;       // hide raw values

        // Adjust the shape of the data labels so they fit the text
        dataLabels.IsResizeShapeToFitText = true; // auto‑fit shape to text
        dataLabels.WidthPixel = 80;                // optional explicit width
        dataLabels.HeightPixel = 30;               // optional explicit height

        // Save the workbook to a file
        workbook.Save("PieChartWithPercentLabels.xlsx");
    }
}
