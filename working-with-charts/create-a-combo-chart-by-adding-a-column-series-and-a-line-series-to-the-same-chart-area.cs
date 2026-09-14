// Title: Create a combo chart with a column series and a line series in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds an Excel workbook, adds a column series for sales data and a line series for profit data to a single chart area using Aspose.Cells. | Show how to assign category labels from a cell range and change the second series type to Line in an Aspose.Cells chart. | Provide the steps to save the workbook as an XLSX file after creating the combo chart with mixed series types. | Explain how to set the chart title and position the chart on the worksheet with Aspose.Cells.
// Common Searches: aspnet aspose.cells create combo chart column and line series example | c# how to add a line series to an existing column chart using Aspose.Cells | set series type to line in Aspose.Cells chart programmatically | assign category axis labels from cells in Aspose.Cells chart c# | save Excel workbook with combo chart using Aspose.Cells .NET
// Tags: Aspose.Cells add column series to chart | Aspose.Cells change series type to line | Aspose.Cells set category data range | Aspose.Cells create combo chart | Aspose.Cells save workbook as xlsx

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a new workbook, fills it with month, sales, and profit data, adds a combo chart positioned on the sheet, defines a column series for sales and a line series for profit, assigns category labels, sets a chart title, and saves the file as ComboChart.xlsx.
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] sales = { 12000, 15000, 13000, 17000, 16000 };
            double[] profit = { 3000, 3500, 3200, 4000, 3800 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
                sheet.Cells[i + 1, 2].PutValue(profit[i]);  // Column C
            }

            // Add a combo chart (Column + Line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Sales and Profit";

            // Define category (X‑axis) labels
            chart.NSeries.CategoryData = "A2:A6";

            // Add column series (Sales)
            int colSeriesIdx = chart.NSeries.Add("B2:B6", true);
            chart.NSeries[colSeriesIdx].Name = "Sales";

            // Add line series (Profit) and set its type to Line
            int lineSeriesIdx = chart.NSeries.Add("C2:C6", true);
            chart.NSeries[lineSeriesIdx].Name = "Profit";
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // Save the workbook
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComboChart.xlsx");
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
