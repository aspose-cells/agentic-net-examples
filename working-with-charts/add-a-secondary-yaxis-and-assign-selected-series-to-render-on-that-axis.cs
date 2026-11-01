// Title: Create an Excel column chart with a secondary Y‑axis and assign a specific series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a workbook that contains a column chart where the second data series is plotted on a secondary Y‑axis with Aspose.Cells in C#. | Demonstrate how to enable the PlotOnSecondAxis property for an NSeries to display that series on the chart's secondary axis using Aspose.Cells. | Write C# code that builds an Excel file with two series, assigns one to the primary axis and the other to a secondary axis, and saves the file.
// Common Searches: how to add a secondary y axis to a column chart using Aspose.Cells C# | Aspose.Cells plot second series on secondary axis example | C# Aspose.Cells dual axis chart tutorial | set PlotOnSecondAxis property Aspose.Cells chart series | create Excel chart with primary and secondary y axes in .NET
// Tags: Aspose.Cells dual axis column chart | C# PlotOnSecondAxis NSeries | Aspose.Cells create chart with secondary Y axis | Excel column chart secondary axis Aspose.Cells | Aspose.Cells chart series assignment

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a new workbook, fills cells with month names and two data series, adds a column chart, adds both series to the chart, sets the second series to plot on the secondary Y‑axis via the PlotOnSecondAxis property, and saves the workbook as ChartWithSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("PrimarySeries");
            sheet.Cells["C1"].PutValue("SecondarySeries");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(200);
            sheet.Cells["C4"].PutValue(300);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (primary Y‑axis)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "PrimarySeries";

            // Add the second series (secondary Y‑axis)
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "SecondarySeries";

            // Plot the second series on the secondary Y‑axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Save the workbook
            workbook.Save("ChartWithSecondaryAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
