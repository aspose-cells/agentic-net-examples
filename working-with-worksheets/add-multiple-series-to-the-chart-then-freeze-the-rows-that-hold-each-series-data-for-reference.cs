// Title: Create a column chart with three data series and freeze the source rows using Aspose.Cells for .NET (C#)
// Common Searches: how to add three series to a column chart with Aspose.Cells in C# | Aspose.Cells C# freeze top rows after creating a chart | set chart series names from header cells using Aspose.Cells .NET | populate worksheet data and bind it to a chart with Aspose.Cells | freeze first five rows in an Excel worksheet using Aspose.Cells C#
// Tags: add three series Aspose.Cells column chart | freeze first five rows Aspose.Cells worksheet | bind series to cell ranges Aspose.Cells | use header cells as series names Aspose.Cells | populate chart data programmatically Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSeriesFreezeDemo
{
    // The example creates a new workbook, fills rows 2‑5 with category labels and three numeric series, adds a column chart, binds each series to its data range, sets the X‑axis categories and series names from the header cells, freezes the first five rows to keep the source data visible while scrolling, and saves the file as MultipleSeriesWithFreeze.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Populate sample data for multiple series
                // ------------------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");
                sheet.Cells["D1"].PutValue("Series 3");

                // Data rows (rows 2‑5)
                string[] categories = { "Q1", "Q2", "Q3", "Q4" };
                int[,] values = {
                    { 120, 150, 180 },
                    { 130, 160, 190 },
                    { 140, 170, 200 },
                    { 150, 180, 210 }
                };

                for (int i = 0; i < categories.Length; i++)
                {
                    int row = i + 2; // Excel rows are 1‑based; Aspose.Cells uses zero‑based indexes internally
                    sheet.Cells[$"A{row}"].PutValue(categories[i]);
                    sheet.Cells[$"B{row}"].PutValue(values[i, 0]);
                    sheet.Cells[$"C{row}"].PutValue(values[i, 1]);
                    sheet.Cells[$"D{row}"].PutValue(values[i, 2]);
                }

                // ------------------------------------------------------------
                // 2. Add a chart to the worksheet
                // ------------------------------------------------------------
                // Chart placed from row 7‑20 and column B‑H (zero‑based indexes)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 1, 19, 7);
                Chart chart = sheet.Charts[chartIndex];

                // ------------------------------------------------------------
                // 3. Add multiple series to the chart
                // ------------------------------------------------------------
                chart.NSeries.Add("=Sheet1!$B$2:$B$5", true); // Series 1
                chart.NSeries.Add("=Sheet1!$C$2:$C$5", true); // Series 2
                chart.NSeries.Add("=Sheet1!$D$2:$D$5", true); // Series 3

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

                // Use header cells as series names
                chart.NSeries[0].Name = "=Sheet1!$B$1";
                chart.NSeries[1].Name = "=Sheet1!$C$1";
                chart.NSeries[2].Name = "=Sheet1!$D$1";

                // ------------------------------------------------------------
                // 4. Freeze rows that contain the series data (rows 2‑5)
                // ------------------------------------------------------------
                // Freeze the first 5 rows (rows 1‑5) while scrolling.
                // Parameters: row index, column index, number of rows to freeze, number of columns to freeze
                sheet.FreezePanes(5, 0, 5, 0);

                // ------------------------------------------------------------
                // 5. Save the workbook
                // ------------------------------------------------------------
                string outputPath = "MultipleSeriesWithFreeze.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output full exception details for debugging
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
