// Title: How to add a second data series to an Aspose.Cells column chart from a worksheet range in C#
// Common Searches: Aspose.Cells C# add another series to an existing column chart using cell range | set category axis and multiple series for a column chart with Aspose.Cells .NET | add series to chart from worksheet range Aspose.Cells C# example | populate Aspose.Cells chart series with data from cells B2:B6 and C2:C6 | create column chart with two series in Aspose.Cells using NSeries.Add
// Tags: NSeries.Add range source Aspose.Cells | assign chart category range Aspose.Cells | column chart series addition Aspose.Cells | export workbook with chart to XLSX Aspose.Cells | C# Aspose.Cells chart series from worksheet cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AddSeriesToChartExample
{
    // The example creates a workbook, fills cells A1:C6 with month and product sales data, adds a column chart, sets the X‑axis category range to A2:A6, adds two series using the ranges B2:B6 and C2:C6 (named "Product A" and "Product B"), and saves the result as ChartWithTwoSeries.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Populate sample data.
                // Column A – Category (e.g., months)
                // Column B – Sales for Product A
                // Column C – Sales for Product B
                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] productA = { 1200, 1500, 1800, 1300, 1700 };
                double[] productB = { 1000, 1400, 1600, 1100, 1500 };

                cells["A1"].PutValue("Month");
                cells["B1"].PutValue("Product A");
                cells["C1"].PutValue("Product B");

                for (int i = 0; i < categories.Length; i++)
                {
                    cells[$"A{i + 2}"].PutValue(categories[i]);
                    cells[$"B{i + 2}"].PutValue(productA[i]);
                    cells[$"C{i + 2}"].PutValue(productB[i]);
                }

                // 3. Add a chart to the worksheet.
                // Parameters: chart type, top row, left column, bottom row, right column.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // 4. Set the category (X‑axis) data for the chart.
                // This tells the chart which cells contain the category labels.
                chart.NSeries.CategoryData = "A2:A6";

                // 5. Add the first series using a worksheet range as the data source.
                // The range "B2:B6" contains the values for Product A.
                // The second parameter (true) indicates that the data is plotted by column.
                chart.NSeries.Add("B2:B6", true);
                // Optionally set a name for the series.
                chart.NSeries[0].Name = "Product A";

                // 6. Add a new series to the same chart using another worksheet range.
                // Here we use the SeriesCollection.Add method, which also accepts a range string.
                // The added series will use the same category data already defined.
                SeriesCollection seriesCollection = chart.NSeries;
                // Add a series for Product B (range C2:C6). The boolean 'true' means vertical plotting.
                seriesCollection.Add("C2:C6", true);
                // The newly added series is at index 1.
                seriesCollection[1].Name = "Product B";

                // 7. Save the workbook to a file.
                string outputPath = "ChartWithTwoSeries.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
