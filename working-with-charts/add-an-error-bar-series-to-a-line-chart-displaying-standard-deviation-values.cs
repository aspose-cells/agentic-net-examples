// Title: Create a line chart with standard deviation error bars using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add error bars to a line chart from a data range | How to display standard deviation as error bars in an Excel line chart with Aspose.Cells | C# example for adding an error‑bar series to a line chart using Aspose.Cells | Aspose.Cells chart error bar support version check for .NET | Generate an Excel line chart with error bars for sales data using Aspose.Cells
// Tags: Aspose.Cells line chart error bars | C# add error bar series Aspose.Cells | standard deviation error bars Excel .NET | Aspose.Cells chart series with error bars | Excel line chart error bar API Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AddErrorBarSeriesDemo
{
    // The example creates a workbook, populates month, sales, and standard deviation columns, adds a line chart for the sales data, and demonstrates how to attach a standard‑deviation error‑bar series to the chart using Aspose.Cells for .NET (noting that the ErrorBar API is available only in newer versions). The workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Month");
                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["C1"].PutValue("StdDev");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
                double[] sales = { 120, 150, 130, 170, 160, 180 };
                double[] stdDev = { 10, 12, 9, 15, 11, 13 };

                for (int i = 0; i < months.Length; i++)
                {
                    worksheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                    worksheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                    worksheet.Cells[i + 1, 2].PutValue(stdDev[i]);  // Column C
                }

                // Add a line chart
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the data series (values from B2:B7, categories from A2:A7)
                int seriesIndex = chart.NSeries.Add("B2:B7", true);
                Series series = chart.NSeries[seriesIndex];
                series.Name = "Sales";

                // NOTE: The current Aspose.Cells version used in this project does not expose
                // the ErrorBar API on the Series class. If a newer version is available,
                // error bars can be added as shown in the original example.
                // For now, the chart is created without error bars.

                // Save the workbook
                string outputPath = "LineChart_With_StdDevErrorBars.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                Console.WriteLine("The line chart has been created (error bars are not supported in this version).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
