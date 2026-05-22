using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    class ComboChartExample
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
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["C1"].PutValue("Profit");

                string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
                int[] sales = { 100, 150, 130, 170, 160 };
                int[] profit = { 30, 45, 35, 50, 40 };

                for (int i = 0; i < months.Length; i++)
                {
                    sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                    sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
                    sheet.Cells[i + 2, 2].PutValue(profit[i]);  // Column C
                }

                // Add a chart (initially a Column chart) to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Add the first series (Sales) – will stay as Column
                chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);

                // Add the second series (Profit) – will be changed to Line
                chart.NSeries.Add("=Sheet1!$C$2:$C$6", true);

                // Set the category (X‑axis) data for both series
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

                // Change the type of the second series to Line to create a combo chart
                chart.NSeries[0].Type = ChartType.Column; // optional, explicit
                chart.NSeries[1].Type = ChartType.Line;

                // Optionally place the line series on a secondary axis
                // Note: IsOnSecondaryAxis property may not be available in older versions.
                // If needed, uncomment the following line after confirming the API support.
                // chart.NSeries[1].IsOnSecondaryAxis = true;

                // Define output file path
                string outputPath = "ComboChart.xlsx";

                // Save the workbook with the combo chart
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}