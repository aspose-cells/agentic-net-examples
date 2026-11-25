// Title: How to place column chart data labels inside the bar ends and avoid overlap using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set column chart data label position to InsideEnd | prevent overlapping data labels in a column chart with Aspose.Cells .NET | configure data label font size and bold style for Aspose.Cells column charts | add a column chart with data labels positioned at the inside end of bars using Aspose.Cells | Aspose.Cells chart label IsNeverOverlap property example in C#
// Tags: Aspose.Cells column chart data label InsideEnd | Aspose.Cells prevent overlapping chart labels | Aspose.Cells set data label font properties C# | Aspose.Cells column chart creation .NET | Aspose.Cells chart label positioning

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart from sample data, enables data labels, positions them inside the end of each bar, disables label overlap with IsNeverOverlap, and applies a larger bold font before saving the file as an XLSX document.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the column chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Configure data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;               // Show the numeric value
        dataLabels.Position = LabelPositionType.InsideEnd; // Position inside the end of each bar
        dataLabels.IsNeverOverlap = true;          // Prevent overlapping of labels

        // Optional: make the labels a bit larger for readability
        dataLabels.Font.Size = 10;
        dataLabels.Font.IsBold = true;

        // Save the workbook to a file
        string outputPath = "ColumnChartDataLabelsInsideEnd.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
        Console.WriteLine($"Workbook saved to '{outputPath}'.");
    }
}
