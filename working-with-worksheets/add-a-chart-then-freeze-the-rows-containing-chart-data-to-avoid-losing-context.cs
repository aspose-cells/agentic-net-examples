// Title: Create a column chart and freeze the data rows in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that inserts a column chart from a specified range and then freezes the rows containing that data. | Show how to call the FreezePanes method after adding a chart to keep the data rows locked while scrolling in an Aspose.Cells workbook. | Adapt the example to position the chart above the dataset and freeze only the header row using the Aspose.Cells C# API.
// Common Searches: Aspose.Cells C# add column chart and freeze specific rows | how to use FreezePanes after creating a chart with Aspose.Cells .NET | freeze rows that contain chart data in Excel using Aspose.Cells C# example | set FreezePanes for rows 1-6 after inserting a chart Aspose.Cells | create column chart from range and lock rows in Aspose.Cells workbook
// Tags: Aspose.Cells create column chart C# | Aspose.Cells FreezePanes rows C# | Excel column chart with frozen data rows Aspose.Cells | Aspose.Cells chart data range to workbook | C# lock rows after chart insertion Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a new workbook, fills month and sales data, adds a column chart based on the sales range, sets a chart title, freezes rows 1‑6 that hold the chart data using FreezePanes, and saves the file as ChartWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 1200, 1500, 1800, 1300, 1700 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A (Month)
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B (Sales)
            }

            // Add a column chart positioned below the data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series
            chart.NSeries.Add("B2:B6", true);
            // Category data is inferred from the first column; explicit setting is optional

            // Set chart title
            chart.Title.Text = "Monthly Sales";

            // Freeze the rows that contain the chart data (rows 1‑6)
            // FreezePanes(row, column, totalRows, totalColumns)
            sheet.FreezePanes(6, 0, 6, 0);

            // Define output file path
            string outputPath = "ChartWithFrozenRows.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
