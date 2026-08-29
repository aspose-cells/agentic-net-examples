// Title: Create a clustered bar chart with a secondary Y‑axis for revenue using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a clustered bar chart from worksheet data and plot the revenue series on a secondary Y‑axis with Aspose.Cells in C#. | Configure the secondary value axis title, minimum, maximum, and major unit for a bar chart using Aspose.Cells. | Save the workbook containing the chart with a secondary axis to an Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells how to plot a series on secondary axis in a bar chart C# | C# code for adding secondary Y axis to clustered bar chart with Aspose.Cells | setting secondary axis title and range for Excel chart using Aspose.Cells .NET | example of dual‑axis bar chart with revenue and units sold in Aspose.Cells
// Tags: Aspose.Cells add secondary axis to bar chart | C# configure secondary value axis Aspose.Cells | Aspose.Cells plot series on second Y axis | dual‑axis clustered bar chart Aspose.Cells | set secondary axis limits Excel Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, fills it with month, units sold, and revenue data, adds a clustered bar chart, assigns the revenue series to a secondary Y‑axis, customizes the secondary axis title and scaling, and saves the file as BarChartWithSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: months, units sold, and revenue
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B1"].PutValue("Units Sold");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(130);

        sheet.Cells["C1"].PutValue("Revenue");
        sheet.Cells["C2"].PutValue(3000);
        sheet.Cells["C3"].PutValue(4500);
        sheet.Cells["C4"].PutValue(3900);

        // Add a clustered bar chart
        int chartIndex = sheet.Charts.Add(ChartType.Bar, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Add two series: Units Sold (primary axis) and Revenue (secondary axis)
        chart.NSeries.Add("B2:B4", true);   // Units Sold
        chart.NSeries.Add("C2:C4", true);   // Revenue
        chart.NSeries.CategoryData = "A2:A4";

        // Assign the revenue series to the secondary Y axis
        chart.NSeries[1].PlotOnSecondAxis = true;

        // Optional: customize the secondary value axis
        Axis secondaryAxis = chart.SecondValueAxis;
        secondaryAxis.Title.Text = "Revenue ($)";
        secondaryAxis.MinValue = 0;
        secondaryAxis.MaxValue = 6000;
        secondaryAxis.MajorUnit = 1000;

        // Save the workbook
        workbook.Save("BarChartWithSecondaryAxis.xlsx");
    }
}
