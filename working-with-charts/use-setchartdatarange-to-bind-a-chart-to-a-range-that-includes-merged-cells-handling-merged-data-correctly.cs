// Title: Bind an Aspose.Cells column chart to a merged‑cell range using SetChartDataRange in C#
// Common Searches: aspnet aspose.cells bind chart to merged cells range c# example | c# setchartdatarange with merged category cells Aspose.Cells | how to create column chart from merged cells using Aspose.Cells | Aspose.Cells chart series from merged cells vertical orientation c# | setchartdatarange vertical series merged cells Aspose.Cells tutorial
// Tags: SetChartDataRange merged cell range | Aspose.Cells column chart merged categories | manual NSeries definition C# | Excel chart merged category cells | bind chart to merged cells Aspose.Cells | vertical series binding Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartMergedCells
{
    // The example creates a workbook, merges cells in column A to represent grouped categories, adds corresponding numeric values, inserts a column chart, binds the chart to the range A1:B5 with SetChartDataRange (vertical series), clears automatic series generation, manually adds the value series and sets X‑values to the merged category range, then saves the file as ChartWithMergedCells.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data.
                // Column A will contain categories, some of which are merged.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // First category spans A2:A3 (merged).
                sheet.Cells["A2"].PutValue("Group 1");
                sheet.Cells.Merge(1, 0, 2, 1); // Merge A2:A3

                // Second category spans A4:A5 (merged).
                sheet.Cells["A4"].PutValue("Group 2");
                sheet.Cells.Merge(3, 0, 2, 1); // Merge A4:A5

                // Corresponding values.
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(15);
                sheet.Cells["B4"].PutValue(20);
                sheet.Cells["B5"].PutValue(25);

                // Add a column chart to the worksheet.
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Bind the chart to a range that includes merged cells.
                // The second argument indicates that the data series are plotted vertically.
                chart.SetChartDataRange("A1:B5", true);

                // Disable automatic series generation and define series manually.
                chart.NSeries.Clear();
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].XValues = "A2:A5";

                // Optional: set chart title.
                chart.Title.Text = "Values by Group";

                // Determine output file path.
                string outputPath = "ChartWithMergedCells.xlsx";

                // Ensure the directory exists (if a directory part is present).
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
