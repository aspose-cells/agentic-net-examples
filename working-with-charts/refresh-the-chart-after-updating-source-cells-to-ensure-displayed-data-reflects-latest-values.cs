// Title: How to refresh an Aspose.Cells column chart after modifying source cells in C#
// AI Prompts: Update worksheet cell values and invoke chart.Calculate() to refresh a linked column chart using Aspose.Cells in C#. | Programmatically bind a column chart to a range and ensure it reflects new data by recalculating the chart with Aspose.Cells. | Save the workbook after changing chart data and calling the chart's Calculate method in a C# application.
// Common Searches: Aspose.Cells C# refresh chart after changing cell values | How to recalculate an Excel chart using Aspose.Cells library | C# example for updating chart data source and calling chart.Calculate | Aspose.Cells column chart not updating after cell edit | Refresh Excel chart programmatically with Aspose.Cells in .NET
// Tags: Aspose.Cells column chart data refresh | recalculate chart after cell update Aspose.Cells | bind chart to range programmatically Aspose.Cells | C# Aspose.Cells chart.Calculate usage | Excel chart refresh using Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example creates a new workbook, adds a column chart bound to cells A2:A4 (categories) and B2:B4 (values), updates the values in B2‑B4, calls chart.Calculate() to refresh the chart, and saves the workbook as ChartRefreshed.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add initial data that will be used by the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart and bind it to the data range
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Update the source cells – this simulates a data change
                worksheet.Cells["B2"].PutValue(15);
                worksheet.Cells["B3"].PutValue(25);
                worksheet.Cells["B4"].PutValue(35);

                // Refresh the chart so it reflects the updated data
                chart.Calculate();

                // Save the workbook
                string outputPath = "ChartRefreshed.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
