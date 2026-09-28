// Title: Create a column chart and a pie chart with separate data ranges on the same worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to insert a Column chart referencing B2:C5 and a Pie chart referencing F2:F5 on the same worksheet, positioning them at H2-O15 and H17-O30 respectively. | Show how to populate two distinct data tables in a workbook and bind each table to a different chart type (Column and Pie) with titles using the Aspose.Cells API. | Generate a complete Aspose.Cells example that creates a workbook, adds sample data, adds multiple charts with custom cell ranges and positions, and saves the file as XLSX.
// Common Searches: Aspose.Cells C# add multiple charts to one worksheet with different data ranges | How to place a column chart and a pie chart at specific cell locations using Aspose.Cells | C# Aspose.Cells example for mixed chart types on the same sheet | Set chart series data from separate columns in Aspose.Cells .NET | Save workbook with several charts using Aspose.Cells for .NET
// Tags: Aspose.Cells add column chart with custom range | Aspose.Cells add pie chart from separate data table | Aspose.Cells position chart using cell coordinates | Aspose.Cells multiple charts on single worksheet | Aspose.Cells save workbook as XLSX with charts | Aspose.Cells bind series data to cells C# | Aspose.Cells mixed chart types example

using System;
using Aspose.Cells;
using Aspose.Cells.Charts; // Required for Chart and ChartType

namespace AsposeCellsMultipleCharts
{
    // Demonstrates how to create a workbook, fill two data tables, add a column chart and a pie chart on the same worksheet with distinct data ranges and cell‑based positions, set titles, and save the file as MultipleCharts.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // -------------------------------------------------
                // Populate data for the first chart (Column Chart)
                // -------------------------------------------------
                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");

                // Sample data
                string[] categories = { "Q1", "Q2", "Q3", "Q4" };
                double[] series1 = { 120, 150, 130, 170 };
                double[] series2 = { 80, 110, 90, 140 };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(categories[i]);   // Column A
                    sheet.Cells[i + 2, 1].PutValue(series1[i]);    // Column B
                    sheet.Cells[i + 2, 2].PutValue(series2[i]);    // Column C
                }

                // -------------------------------------------------
                // Populate data for the second chart (Pie Chart)
                // -------------------------------------------------
                // Header
                sheet.Cells["E1"].PutValue("Product");
                sheet.Cells["F1"].PutValue("Sales");

                // Sample data
                string[] products = { "Product A", "Product B", "Product C", "Product D" };
                double[] sales = { 300, 500, 200, 400 };

                for (int i = 0; i < products.Length; i++)
                {
                    sheet.Cells[i + 2, 4].PutValue(products[i]); // Column E
                    sheet.Cells[i + 2, 5].PutValue(sales[i]);   // Column F
                }

                // -------------------------------------------------
                // Create the first chart (Column Chart) on the same worksheet
                // -------------------------------------------------
                // Position the chart at cells H2 to O15
                int chartIndex1 = sheet.Charts.Add(ChartType.Column, 1, 7, 14, 14);
                Chart chart1 = sheet.Charts[chartIndex1];
                chart1.Title.Text = "Quarterly Sales Comparison";

                // Add series: Series1 and Series2 using the data range B2:C5
                // Category (X) axis uses A2:A5
                chart1.NSeries.Add("B2:C5", true);
                chart1.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // Create the second chart (Pie Chart) on the same worksheet
                // -------------------------------------------------
                // Position the chart at cells H17 to O30
                int chartIndex2 = sheet.Charts.Add(ChartType.Pie, 16, 7, 29, 14);
                Chart chart2 = sheet.Charts[chartIndex2];
                chart2.Title.Text = "Product Sales Distribution";

                // Add series using the data range F2:F5
                // Category (slice names) uses E2:E5
                chart2.NSeries.Add("F2:F5", true);
                chart2.NSeries.CategoryData = "E2:E5";

                // -------------------------------------------------
                // Save the workbook to a file
                // -------------------------------------------------
                workbook.Save("MultipleCharts.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
