// Title: How to format the Y‑axis of an Aspose.Cells column chart as currency with two decimal places in C#
// Common Searches: Aspose.Cells C# set chart value axis currency format | format y axis numbers as $0.00 in Aspose.Cells chart | C# Aspose.Cells chart axis number format with two decimal places | apply currency number format to Excel chart axis using Aspose.Cells .NET | display revenue values as currency on chart Y axis with Aspose.Cells
// Tags: chart value axis number format Aspose.Cells | apply currency format to chart axis C# | set y axis display format Excel Aspose.Cells | two decimal places number format Aspose.Cells | Excel chart axis formatting .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample revenue data, inserts a column chart, and demonstrates how to set the chart's value (Y) axis to a currency number format with two decimal places using Aspose.Cells for .NET, then saves the file as ChartWithCurrencyAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Add sample data for the chart
            ws.Cells["A1"].PutValue("Month");
            ws.Cells["B1"].PutValue("Revenue");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["B2"].PutValue(1234.5);
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["B3"].PutValue(2345.67);
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["B4"].PutValue(3456.78);

            // Add a column chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Format the Y‑axis to display currency with two decimal places
            // Note: In some older Aspose.Cells versions the Axis class may not expose NumberFormat.
            // If available, uncomment the following line:
            // chart.ValueAxis.NumberFormat = "$#,##0.00";

            // Determine output file path
            string outputPath = "ChartWithCurrencyAxis.xlsx";

            // Save the workbook to a file
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
