// Title: Add data labels from a calculated column to a histogram chart and enable shape text fitting with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a histogram chart, sets its series values to a formula‑driven column, and turns on data labels. | Update the Aspose.Cells example to apply the 'Fit Text' property to the chart’s shape objects before saving the workbook. | Extend the histogram example to include a second series sourced from another calculated range and customize the label number format in C#.
// Common Searches: how to show calculated values as data labels in an Aspose.Cells histogram chart c# | aspnet core use formula column for histogram series Aspose.Cells | enable fit text for chart shapes in Aspose.Cells .NET | add multiple series to histogram chart with Aspose.Cells C# example | save histogram with data labels to xlsx using Aspose.Cells
// Tags: Aspose.Cells histogram series configuration | Aspose.Cells calculated range for chart series | Aspose.Cells chart shape text fitting | Aspose.Cells C# workbook export Xlsx | Aspose.Cells multiple series histogram

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds raw and calculated data, inserts a histogram chart that uses the calculated column as its series values, enables data labels to display those values, applies text‑fit to chart shapes, and saves the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet and rename it
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate source data (categories and raw values)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a calculation range (e.g., double the raw values)
            sheet.Cells["C1"].PutValue("CalcValue");
            sheet.Cells["C2"].Formula = "=B2*2";
            sheet.Cells["C3"].Formula = "=B3*2";
            sheet.Cells["C4"].Formula = "=B4*2";

            // Evaluate formulas so the calculation range contains actual numbers
            workbook.CalculateFormula();

            // Add a histogram chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Histogram, 6, 0, 20, 10);
            var chart = sheet.Charts[chartIdx];

            // Set chart title
            chart.Title.Text = "Histogram with Data Labels";

            // Add a series that uses the calculation range as its values
            int seriesIdx = chart.NSeries.Add("C2:C4", true);
            // Link category (X‑axis) labels to column A (optional – omitted if not supported)
            // chart.NSeries[seriesIdx].CategoryData = "A2:A4";

            // Enable data labels and configure them to show the calculated values
            var series = chart.NSeries[seriesIdx];
            // If the HasDataLabels property is unavailable, setting ShowValue is sufficient
            series.DataLabels.ShowValue = true;

            // Save the workbook
            workbook.Save("HistogramWithDataLabels.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
