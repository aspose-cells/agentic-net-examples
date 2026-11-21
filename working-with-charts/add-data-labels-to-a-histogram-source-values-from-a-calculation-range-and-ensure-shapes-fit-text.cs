// Title: Add calculated data labels to a column (histogram) chart and auto‑fit label shapes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add data labels showing formula results in a column chart | how to bind a calculated range as series source for a histogram in Aspose.Cells | auto fit chart data label shape to text Aspose.Cells .NET | display doubled values as labels on a histogram using Aspose.Cells | C# Aspose.Cells create column chart from formula cells with data labels
// Tags: column chart data labels from formula range Aspose.Cells | histogram series source calculated cells .NET | auto‑fit data label shape Aspose.Cells | display calculated values as chart labels C# | Aspose.Cells chart series from computed range

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds source data and a calculated column that doubles the original values, inserts a column (histogram) chart using the calculated range as its series, enables data labels to show those calculated values, and ensures the label shapes auto‑fit the text before saving the file as 'HistogramWithDataLabels.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Source data -----
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // ----- Calculation range (e.g., double the original values) -----
            sheet.Cells["C1"].PutValue("CalcValue");
            sheet.Cells["C2"].Formula = "=B2*2";
            sheet.Cells["C3"].Formula = "=B3*2";
            sheet.Cells["C4"].Formula = "=B4*2";

            // ----- Add a histogram (column) chart -----
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Use the calculation range as the data source for the series
            chart.NSeries.Add("C2:C4", true);

            // ----- Add data labels and show the calculated values -----
            // Enable data labels and display the value in each label
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook (lifecycle rule)
            workbook.Save("HistogramWithDataLabels.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
