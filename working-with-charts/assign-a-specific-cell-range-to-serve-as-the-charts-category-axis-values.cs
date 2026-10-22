// Title: Assign a specific cell range as the category axis labels for a column chart with Aspose.Cells in C#
// AI Prompts: Generate a column chart and set its category axis to the cells A2:A5 using Aspose.Cells NSeries.CategoryData in C#. | Bind worksheet range A2:A5 as category labels for a chart series programmatically with Aspose.Cells for .NET. | Create a workbook, add sample data, and configure the chart’s category data range without using the UI.
// Common Searches: Aspose.Cells C# bind category labels to a chart from a worksheet range | set NSeries.CategoryData property for a column chart in .NET | example code assigning chart category axis values using Aspose.Cells | C# create column chart with custom category axis range Aspose.Cells | Aspose.Cells chart series category data from cells A2:A5
// Tags: Aspose.Cells set chart category data range | C# NSeries.CategoryData usage | column chart category axis from worksheet cells | Aspose.Cells chart series values and categories | programmatic chart axis binding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, fills cells A1:B5 with sample data, adds a column chart, assigns the series values from B2:B5, sets the category axis labels to the range A2:A5 via the NSeries.CategoryData property, and saves the workbook as ChartWithCategoryAxis.xlsx.
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

                // Populate sample data for the chart (optional but makes the chart visible)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Add a column chart to the worksheet (positioned from row 5, column 0 to row 15, column 5)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Add a data series for the chart (values from B2:B5)
                chart.NSeries.Add("B2:B5", true);

                // Assign a specific cell range as the category axis values (A2:A5)
                chart.NSeries.CategoryData = "A2:A5";

                // Define output file path
                string outputPath = "ChartWithCategoryAxis.xlsx";

                // Save the workbook
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
