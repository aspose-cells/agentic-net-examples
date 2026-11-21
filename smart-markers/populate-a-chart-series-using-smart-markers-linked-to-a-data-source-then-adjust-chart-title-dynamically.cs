// Title: Create an Excel column chart with smart‑marker series binding and a dynamic total‑sales title using Aspose.Cells for .NET
// Common Searches: aspnet create column chart using smart markers Aspose.Cells | bind chart series to range with smart marker expression Aspose.Cells C# | calculate sum of column data and set chart title dynamically Aspose.Cells | format Excel chart title as currency in C# Aspose.Cells
// Tags: smart markers bind chart series Aspose.Cells | dynamic chart title from calculated sum Aspose.Cells | column chart from C# array data Aspose.Cells | currency formatted chart title .NET | programmatic Excel chart creation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, adds month and sales data, inserts a column chart, binds the series to the sales range using a smart‑marker expression, computes the total sales, sets the chart title to the formatted sum, and saves the file as SmartMarkerChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it.
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Data";

                // Populate sample data.
                // Column A: Categories (e.g., Months)
                // Column B: Values (e.g., Sales)
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");

                string[] months = { "Jan", "Feb", "Mar", "Apr" };
                double[] sales = { 1200.5, 1500.75, 1100.0, 1800.25 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                    sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
                }

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Link the series to the data source using smart markers.
                chart.NSeries.Add("{=Data!$B$2:$B$5}", true);
                // Category data is automatically inferred when using smart markers,
                // so explicit assignment is not required.

                // Dynamically set the chart title based on the sum of sales.
                double totalSales = 0;
                foreach (double value in sales)
                {
                    totalSales += value;
                }
                chart.Title.Text = $"Total Sales: {totalSales:C}";
                chart.Title.Font.Size = 12;
                chart.Title.Font.IsBold = true;

                // Ensure the output directory exists.
                string outputPath = "SmartMarkerChart.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
