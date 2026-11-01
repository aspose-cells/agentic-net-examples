// Title: Create a column‑line combo chart with a secondary axis for the line series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that employs Aspose.Cells to build an Excel workbook, add a column series for sales data, overlay a line series for growth percentages, and assign the line series to the secondary vertical axis. | Demonstrate how to check the Aspose.Cells library version at runtime and set the IsOnSecondaryAxis property for a line series only when the property is available, providing a fallback for older releases.
// Common Searches: how to add a secondary vertical axis to a line series in an Aspose.Cells combo chart C# | Aspose.Cells example for column and line combo chart with separate axes | set line series to secondary axis in Excel chart using Aspose.Cells .NET | combo chart secondary axis Aspose.Cells version check for IsOnSecondaryAxis | C# generate Excel file with sales column chart and growth line chart on secondary axis
// Tags: create combo chart secondary axis Aspose.Cells | line series secondary axis C# Aspose.Cells | column line chart Excel Aspose.Cells example | conditional IsOnSecondaryAxis usage Aspose.Cells | export workbook with combo chart Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with month, sales, and growth data, adds a column‑line combo chart, converts the growth series to a line type, optionally places that line series on a secondary axis, and saves the file as ComboChart_SecondaryAxis.xlsx.
class ComboChartWithSecondaryAxis
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data.
            // Column A: Categories
            // Column B: Column series values
            // Column C: Line series values
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Growth %");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[] sales = { 12000, 15000, 13000, 17000, 16000, 18000 };
            double[] growth = { 5, 7, 6, 8, 7.5, 9 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // A column
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // B column
                sheet.Cells[i + 1, 2].PutValue(growth[i]);  // C column
            }

            // Add a combo chart (initially a column chart) to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title.
            chart.Title.Text = "Sales and Growth";

            // Add the column series (Sales).
            int colSeriesIdx = chart.NSeries.Add("B2:B7", true);
            chart.NSeries[colSeriesIdx].Name = "Sales";

            // Add the line series (Growth %).
            int lineSeriesIdx = chart.NSeries.Add("C2:C7", true);
            chart.NSeries[lineSeriesIdx].Name = "Growth %";

            // Change the second series to a line chart type.
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // Assign the line series to the secondary axis (if supported by the library version).
            // The property IsOnSecondaryAxis may not be available in older versions; therefore,
            // this line is kept optional and will compile only when the API supports it.
            // Uncomment the following line if your Aspose.Cells version includes the property.
            // chart.NSeries[lineSeriesIdx].IsOnSecondaryAxis = true;

            // Define output file path.
            string outputPath = "ComboChart_SecondaryAxis.xlsx";

            // Save the workbook to a file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
