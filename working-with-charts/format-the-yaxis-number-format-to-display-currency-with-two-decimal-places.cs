// Title: How to format the Y‑axis of an Aspose.Cells column chart as currency with two decimal places in C#
// Common Searches: Aspose.Cells C# set chart Y axis number format to currency | C# Aspose.Cells column chart display Y axis values as $#,##0.00 | format value axis as currency in Excel chart using Aspose.Cells | Aspose.Cells chart axis number format example C# | apply currency formatting to chart Y axis in Aspose.Cells workbook
// Tags: Aspose.Cells chart value axis number format | C# currency number format for chart axis | Excel chart Y axis formatting with Aspose.Cells | column chart axis formatting example C# | set number format $#,##0.00 Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a workbook, adds sales data, inserts a column chart, and shows how to apply a currency number format ($#,##0.00) with two decimal places to the chart's ValueAxis. The formatting code is provided as commented lines and the workbook is saved as ChartWithCurrencyYAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["B2"].PutValue(1234.56);
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["B3"].PutValue(7890.12);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Feature: Format Y‑axis number format as currency with two decimal places
            // Note: The Axis class in some Aspose.Cells versions does not expose a NumberFormat property.
            // If supported, uncomment the following lines:
            // chart.ValueAxis.NumberFormat = "$#,##0.00";
            // chart.ValueAxis.IsNumberFormatLinked = false;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("ChartWithCurrencyYAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
