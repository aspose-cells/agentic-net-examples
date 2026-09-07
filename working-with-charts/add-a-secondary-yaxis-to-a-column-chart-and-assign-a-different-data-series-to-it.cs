// Title: Add a secondary Y‑axis to a clustered column chart and map a separate series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate an Excel workbook with month, sales, and profit data, then create a clustered column chart where the sales series uses the primary Y‑axis and the profit series is plotted on a secondary Y‑axis using Aspose.Cells C# API. | Write C# code that adds a column chart to a worksheet, assigns one data series to the primary axis and another to the secondary axis, and saves the file as an .xlsx. | Produce a script that demonstrates how to enable a secondary Y‑axis for a column chart and set PlotOnSecondAxis = true for a specific series with Aspose.Cells.
// Common Searches: Aspose.Cells how to display two Y axes in a column chart C# | C# create Excel chart with primary and secondary axis using Aspose.Cells | example of PlotOnSecondAxis property in Aspose.Cells chart | dual axis column chart Aspose.Cells .NET tutorial | assign profit series to secondary axis in Aspose.Cells chart
// Tags: Aspose.Cells column chart secondary axis | C# Aspose.Cells dual Y axis chart | PlotOnSecondAxis property Aspose.Cells | Excel chart multiple axes Aspose.Cells .NET | assign series to secondary axis C# Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, fills it with month, sales, and profit data, adds a clustered column chart, plots the Sales series on the primary Y‑axis and the Profit series on a secondary Y‑axis, and saves the workbook as ChartWithSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the chart
            // Category labels
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Primary series data (Sales)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);

            // Secondary series data (Profit)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(55);

            // Add a clustered column chart (use ChartType.Column for compatibility)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sales and Profit";

            // Add the primary series (Sales) – plotted on the primary Y‑axis
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Sales";

            // Add the secondary series (Profit) – plotted on the secondary Y‑axis
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Profit";
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Note: In older Aspose.Cells versions the secondary axis becomes visible automatically
            // when a series is plotted on it, so explicit axis handling is omitted.

            // Save the workbook
            string outputPath = "ChartWithSecondaryAxis.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
