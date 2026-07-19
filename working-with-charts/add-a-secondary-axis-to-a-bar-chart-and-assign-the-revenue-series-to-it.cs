// Title: C# – Add a Secondary Y‑Axis to a Bar Chart using Aspose.Cells
// Description: Creates a workbook, fills it with month, units sold, and revenue data, adds a 2‑D bar chart, assigns the revenue series to a secondary Y‑axis, customizes the axis title and scale, and saves the file as BarChartWithSecondaryAxis.xlsx.
// Keywords: Aspose.Cells secondary axis C# | bar chart secondary Y axis .NET | plot series on secondary axis Aspose.Cells | customize chart axis Aspose.Cells | Aspose.Cells chart scaling
// Common Searches: Aspose.Cells add secondary axis to bar chart | C# plot series on secondary value axis Aspose.Cells | set secondary axis title and range Aspose.Cells | how to use PlotOnSecondAxis in Aspose.Cells | secondary Y axis example Aspose.Cells .NET
// Developer Intent: Generate a bar chart where the revenue series is displayed on a secondary Y‑axis.
// Use Cases: Compare units sold and revenue when the metrics have different scales. | Create financial reports that show sales volume and monetary value side‑by‑side. | Provide a clear axis label and range for revenue to improve chart readability.
// AI Prompts: Write C# code that adds a secondary Y‑axis to a bar chart and assigns a specific series to it using Aspose.Cells. | Explain how to configure the secondary axis title, minimum, maximum, and major unit after creating a chart with Aspose.Cells. | Show how to plot multiple series on primary and secondary axes in a single Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisDemo
{
    // Creates a workbook, fills it with month, units sold, and revenue data, adds a 2‑D bar chart, assigns the revenue series to a secondary Y‑axis, customizes the axis title and scale, and saves the file as BarChartWithSecondaryAxis.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Category
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Column B: Units Sold
            sheet.Cells["B1"].PutValue("Units Sold");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Column C: Revenue
            sheet.Cells["C1"].PutValue("Revenue");
            sheet.Cells["C2"].PutValue(3000);
            sheet.Cells["C3"].PutValue(3750);
            sheet.Cells["C4"].PutValue(4500);

            // Add a 2‑D bar chart
            int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add the first series (Units Sold) – primary axis
            chart.NSeries.Add("B2:B4", true);
            // Add the second series (Revenue) – will be plotted on secondary axis
            chart.NSeries.Add("C2:C4", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Assign the revenue series (index 1) to the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize the secondary value axis (e.g., title and range)
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Revenue (USD)";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // Save the workbook
            workbook.Save("BarChartWithSecondaryAxis.xlsx");
        }
    }
}
