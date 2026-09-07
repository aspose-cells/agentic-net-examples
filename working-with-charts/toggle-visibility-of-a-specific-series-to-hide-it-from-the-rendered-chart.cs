// Title: Hide a specific data series in an Excel column chart using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that creates a column chart, adds two series, and sets the first series IsFiltered = true to hide it. | Write a C# snippet that reads the IsFiltered flag of each chart series and logs the visibility status before saving the workbook with Aspose.Cells. | Provide an example that toggles a chart series' visibility on demand and saves the workbook as an .xlsx file using Aspose.Cells.
// Common Searches: aspnet hide first series in column chart using Aspose.Cells | how to use IsFiltered property to hide chart series in Aspose.Cells C# | programmatically filter out a series from an Excel chart with Aspose.Cells .NET | toggle visibility of chart series at runtime Aspose.Cells | Aspose.Cells column chart series visibility example C#
// Tags: Aspose.Cells IsFiltered chart series | hide column chart series C# | toggle chart series visibility Aspose.Cells | filter out series Excel chart .NET | chart series visibility control Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesVisibilityDemo
{
    // The sample creates a workbook, fills it with sample data, builds a column chart with two series, hides the first series by setting its IsFiltered property to true, prints each series' visibility flag, and saves the file as SeriesVisibilityDemo.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for two series
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                // Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set data ranges for the two series
                chart.NSeries.Add("B2:B5", true); // Series1
                chart.NSeries.Add("C2:C5", true); // Series2
                chart.NSeries.CategoryData = "A2:A5";

                // Hide the first series (Series1) by setting IsFiltered = true
                if (chart.NSeries.Count > 0)
                {
                    chart.NSeries[0].IsFiltered = true;
                }

                // Output visibility status safely
                Console.WriteLine("Series1 IsFiltered = " +
                                  (chart.NSeries.Count > 0 ? chart.NSeries[0].IsFiltered.ToString() : "N/A"));
                Console.WriteLine("Series2 IsFiltered = " +
                                  (chart.NSeries.Count > 1 ? chart.NSeries[1].IsFiltered.ToString() : "N/A"));

                // Save the workbook
                string outputPath = "SeriesVisibilityDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
