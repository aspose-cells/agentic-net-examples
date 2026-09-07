// Title: How to add median data labels to a Box‑and‑Whisker chart from a linked cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a Box‑and‑Whisker chart, links its data labels to a separate column, and displays the labels inside each box using Aspose.Cells. | Write C# code that activates ShowValue and ShowCellRange on a chart series, assigns LinkedSource to a cell range, and displays custom summary values in the chart labels using Aspose.Cells. | Show how to configure the label placement to appear within each box of a Box‑and‑Whisker chart and export the workbook to an XLSX file with Aspose.Cells.
// Common Searches: asp.net aspose.cells link box whisker chart data labels to cell range | c# add median labels to box and whisker chart using aspose cells | how to show statistical summary values as data labels in aspose cells chart | set data label position inside box whisker series aspose.cells .net | asp.net create box whisker chart with custom data labels from another column
// Tags: box-whisker chart linked data label range | aspnet aspose.cells median label association | c# chart series showcellrange property | inside label positioning for box-whisker chart | aspose.cells create box-whisker chart xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates using Aspose.Cells for .NET to build a Box‑and‑Whisker chart, populate categories and raw values, add a median column, enable data labels, link them to the median range, position the labels inside the boxes, and save the workbook as an XLSX file.
class BoxWhiskerDataLabelsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the box‑and‑whisker chart
        // Column A – categories
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q1");
        sheet.Cells["A4"].PutValue("Q1");
        sheet.Cells["A5"].PutValue("Q2");
        sheet.Cells["A6"].PutValue("Q2");
        sheet.Cells["A7"].PutValue("Q2");

        // Column B – raw values used to build the box‑and‑whisker
        sheet.Cells["B1"].PutValue("Values");
        sheet.Cells["B2"].PutValue(15);
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(20);
        sheet.Cells["B6"].PutValue(35);
        sheet.Cells["B7"].PutValue(40);

        // Column C – statistical summary that will be shown in data labels
        // (e.g., median values for each category)
        sheet.Cells["C1"].PutValue("Median");
        sheet.Cells["C2"].PutValue(25); // median for Q1
        sheet.Cells["C5"].PutValue(30); // median for Q2

        // Add a box‑and‑whisker chart
        int chartIndex = sheet.Charts.Add(ChartType.BoxWhisker, 5, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart (raw values)
        chart.SetChartDataRange("B1:B7", true);
        chart.NSeries.CategoryData = "A2:A7";

        // Access the first (and only) series
        Series series = chart.NSeries[0];

        // Enable data labels and bind them to the statistical summary range
        series.DataLabels.ShowValue = true;          // show the value (median) in the label
        series.DataLabels.ShowCellRange = true;     // indicate that the label comes from a cell range
        series.DataLabels.LinkedSource = "C2:C5";   // range containing the summary values

        // Optional: position the data labels inside the box
        series.DataLabels.Position = LabelPositionType.InsideBase;

        // Save the workbook
        workbook.Save("BoxWhisker_WithDataLabels.xlsx");
    }
}
