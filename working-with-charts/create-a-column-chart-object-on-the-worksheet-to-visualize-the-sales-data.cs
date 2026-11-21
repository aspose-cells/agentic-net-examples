// Title: Generate a column chart for monthly sales data in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: how to create a column chart from a range of cells using Aspose.Cells C# | Aspose.Cells C# add chart title and series name to an Excel column chart | save Excel workbook with a column chart to XLSX using Aspose.Cells .NET | populate worksheet with month and sales values and generate a chart Aspose.Cells | set chart categories and values programmatically with Aspose.Cells C#
// Tags: Aspose.Cells create column chart | Aspose.Cells set chart title | Aspose.Cells add series from range | Aspose.Cells populate worksheet with data | Aspose.Cells save workbook as xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, writes month names and sales figures into columns A and B, adds a column chart covering rows 5‑25, sets the chart title to "Monthly Sales", defines the series values from B2:B6 with categories from A2:A6, names the series "Sales", and saves the file as SalesColumnChart.xlsx.
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
                sheet.Name = "SalesData";

                // Populate sales data (Month and Sales)
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["B1"].PutValue("Sales");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] sales = { 15000, 18000, 12000, 21000, 17000 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A (Month)
                    sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B (Sales)
                }

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Monthly Sales";

                // Add a series (values from B2:B6)
                int seriesIndex = chart.NSeries.Add("B2:B6", true);
                // Set categories for the whole NSeries collection
                chart.NSeries.CategoryData = "A2:A6";

                // Optional: Customize the series name
                chart.NSeries[seriesIndex].Name = "Sales";

                // Determine output file path
                string outputFile = "SalesColumnChart.xlsx";
                string outputPath = Path.GetFullPath(outputFile);

                // Ensure the directory exists (if any)
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
