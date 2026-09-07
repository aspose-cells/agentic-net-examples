// Title: Create a column chart from worksheet data and save the workbook as an XLSX file using Aspose.Cells for .NET
// AI Prompts: Write C# code that builds a new workbook, populates cells A1:B5 with category and value data, adds a column chart linked to that range, sets a custom chart title and series name, and saves the workbook as an XLSX file with Aspose.Cells. | Refactor the example to accept the output file path as a parameter, switch the chart type to a line chart, and update the chart title accordingly using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# add column chart from cell range and export to XLSX | how to set chart title and series name in Aspose.Cells .NET | save workbook containing a chart as XLSX using Aspose.Cells C# | Aspose.Cells create chart, populate data, and specify output filename
// Tags: Aspose.Cells create column chart C# | Aspose.Cells save workbook as XLSX | Aspose.Cells set chart title C# | Aspose.Cells define NSeries data range | Aspose.Cells populate worksheet for chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart and ChartType

namespace AsposeCellsChartExample
{
    // The program creates a new workbook, fills cells A1:B5 with sample categories and values, adds a column chart referencing the data, customizes the chart title and series name, and saves the workbook as GeneratedChart.xlsx.
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

                // Populate sample data for the chart
                // A1:A5 - Categories, B1:B5 - Values
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Monthly Sales";

                // Add a series (values) – the second parameter indicates that the data is vertical
                int seriesIndex = chart.NSeries.Add("B2:B5", true);
                // Optionally set the category (X) axis data if the API version supports it
                // If CategoryData property is unavailable, Aspose.Cells will use the first column as categories by default
                // chart.NSeries[seriesIndex].CategoryData = "A2:A5";

                // Optionally customize the series name
                chart.NSeries[seriesIndex].Name = "Sales";

                // Save the workbook as an XLSX file
                string outputPath = "GeneratedChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
