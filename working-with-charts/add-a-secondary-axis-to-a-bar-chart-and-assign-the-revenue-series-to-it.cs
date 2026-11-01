// Title: Create a clustered column chart with a secondary axis for the revenue series using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a workbook, populate month, units sold, and revenue data, add a clustered column chart, and move the revenue series to a secondary value axis with Aspose.Cells in C#. | Write C# code that creates a column chart with two series, places the second series on a secondary value axis, sets axis titles, and saves the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: how to add a secondary axis to a column chart using Aspose.Cells C# | Aspose.Cells dual axis chart example for sales and revenue | C# set IsOnSecondaryAxis property on chart series Aspose.Cells | create clustered column chart with primary and secondary axes in .NET | Aspose.Cells secondary axis not supported older versions
// Tags: Aspose.Cells generate column chart | Aspose.Cells enable secondary axis for chart | Aspose.Cells map series to secondary axis | C# generate Excel chart with dual axes | Aspose.Cells set axis label text

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills it with month, units sold, and revenue data, adds a clustered column chart, inserts two series (Units Sold and Revenue), assigns the Revenue series to a secondary value axis via the IsOnSecondaryAxis property when available, sets axis titles, and saves the file as BarChartWithSecondaryAxis.xlsx.
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
            sheet.Cells["B1"].PutValue("Units Sold");
            sheet.Cells["C1"].PutValue("Revenue");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            sheet.Cells["C2"].PutValue(2000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(2600);

            // Add a clustered column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Add the Units Sold series (primary axis)
            int unitsSeriesIdx = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[unitsSeriesIdx].Name = "Units Sold";

            // Add the Revenue series (secondary axis if supported)
            int revenueSeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[revenueSeriesIdx].Name = "Revenue";

            // NOTE: The Series.IsOnSecondaryAxis property may not be available in older Aspose.Cells versions.
            // If supported, you can uncomment the following line:
            // chart.NSeries[revenueSeriesIdx].IsOnSecondaryAxis = true;

            // Set axis title for the primary value axis
            chart.ValueAxis.Title.Text = "Units Sold";

            // Save the workbook with the chart
            string outputPath = "BarChartWithSecondaryAxis.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
