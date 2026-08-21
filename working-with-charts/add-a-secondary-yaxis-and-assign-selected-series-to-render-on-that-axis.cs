// Title: Aspose.Cells for .NET – Add a Secondary Y‑Axis to a Column Chart and Plot a Selected Series
// Description: Creates a workbook, inserts category, revenue, and units‑sold data, builds a column chart, adds two series, assigns the second series to a secondary Y‑axis, customizes its title and scale, and saves the file as ChartWithSecondaryYAxis.xlsx.
// Keywords: Aspose.Cells secondary axis | C# chart secondary Y axis | plot series on second axis Aspose | customize chart axes .NET | column chart multiple Y axes | Aspose.Cells chart example
// Common Searches: Aspose.Cells add secondary Y axis to chart | C# plot series on second axis in Excel | set secondary axis title Aspose.Cells | configure secondary value axis range .NET | column chart with two Y axes Aspose
// Developer Intent: Generate a column chart where one data series uses the primary Y‑axis and another uses a secondary Y‑axis, with optional axis formatting.
// Use Cases: Display revenue (currency) alongside units sold (quantity) in a single visual with distinct scales. | Combine financial and operational metrics in automated Excel reports for executive dashboards. | Create multi‑axis charts for scientific data where measurements have different units.
// AI Prompts: Write C# code using Aspose.Cells to add a secondary Y‑axis to an existing column chart and assign a specific series to it. | Show how to set the title, minimum, maximum, and major unit of the secondary axis in an Aspose.Cells chart. | Explain how to hide the secondary axis while still rendering its series in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryYAxisDemo
{
    // Creates a workbook, inserts category, revenue, and units‑sold data, builds a column chart, adds two series, assigns the second series to a secondary Y‑axis, customizes its title and scale, and saves the file as ChartWithSecondaryYAxis.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A – Category (X‑axis)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Column B – First series (plotted on primary Y‑axis)
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["B2"].PutValue(12000);
            sheet.Cells["B3"].PutValue(15000);
            sheet.Cells["B4"].PutValue(18000);

            // Column C – Second series (will be plotted on secondary Y‑axis)
            sheet.Cells["C1"].PutValue("Units Sold");
            sheet.Cells["C2"].PutValue(300);
            sheet.Cells["C3"].PutValue(450);
            sheet.Cells["C4"].PutValue(600);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (Revenue) – uses column B values
            chart.NSeries.Add("B2:B4", true);
            // Add the second series (Units Sold) – uses column C values
            chart.NSeries.Add("C2:C4", true);

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // Assign the second series to the secondary Y‑axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize the secondary Y‑axis appearance
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Units Sold";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 800;
            secondaryAxis.MajorUnit = 200;
            secondaryAxis.IsVisible = true; // ensure it is shown

            // Save the workbook
            workbook.Save("ChartWithSecondaryYAxis.xlsx");
        }
    }
}
