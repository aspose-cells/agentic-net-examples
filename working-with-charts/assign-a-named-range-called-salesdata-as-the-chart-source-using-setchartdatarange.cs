// Title: How to bind a named range to a column chart using SetChartDataRange in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart data source to a named range | Using SetChartDataRange with a workbook named range in .NET | Create a named range and attach it to a column chart in Aspose.Cells | Example of chart.NSeries.CategoryData with a named range in C# | Bind SalesData named range to a chart series using Aspose.Cells API
// Tags: Aspose.Cells SetChartDataRange named range | C# column chart from named range | Aspose.Cells create named range workbook | chart series data binding Aspose.Cells | Excel chart data source named range .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartNamedRangeDemo
{
    // The example creates a new workbook, defines a named range called SalesData that references cells B2:B5, adds a column chart, sets the chart's data source to the SalesData named range with SetChartDataRange, assigns category labels from A2:A5, optionally names the series, and saves the file as ChartWithNamedRange.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Sales");

                string[] categories = { "North", "South", "East", "West" };
                double[] sales = { 12000, 15000, 9000, 13000 };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(categories[i]); // Column A
                    sheet.Cells[i + 2, 1].PutValue(sales[i]);    // Column B
                }

                // Add a named range "SalesData" that refers to the sales values (B2:B5)
                // The Add method with a single argument creates the name; the RefersTo property sets the range.
                int nameIndex = workbook.Worksheets.Names.Add("SalesData");
                Name salesDataName = workbook.Worksheets.Names[nameIndex];
                salesDataName.RefersTo = "=Sheet1!$B$2:$B$5";

                // Verify that the named range was created
                if (salesDataName == null)
                {
                    Console.WriteLine("Failed to create the named range 'SalesData'.");
                    return;
                }
                Console.WriteLine($"Named range 'SalesData' created with RefersTo: {salesDataName.RefersTo}");

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Use the named range as the chart source
                chart.SetChartDataRange("SalesData", true);

                // Set category (X‑axis) data so the chart shows the category names
                chart.NSeries.CategoryData = "A2:A5";

                // Optionally give the series a name
                if (chart.NSeries.Count > 0)
                {
                    chart.NSeries[0].Name = "Quarterly Sales";
                }

                // Save the workbook
                string outputPath = "ChartWithNamedRange.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
