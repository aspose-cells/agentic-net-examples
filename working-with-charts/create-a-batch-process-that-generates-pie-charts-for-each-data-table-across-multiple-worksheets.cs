// Title: Create pie charts for each Excel table on every worksheet with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook, iterates through all worksheets and their ListObjects, and inserts a pie chart for each table using Aspose.Cells. | Generate a pie chart where the series is bound to the table's data range, the chart title comes from the table's display name, and data labels show percentages and category names. | Place each pie chart in a specific cell block (e.g., rows 5‑20, columns A‑K) and save the workbook after processing all tables.
// Common Searches: aspnet add pie chart to every table in an Excel workbook using Aspose.Cells | c# generate pie charts for all ListObjects across multiple worksheets | how to bind Aspose.Cells pie chart series to a table data range | batch create pie charts in a workbook and set chart titles from table names with Aspose.Cells
// Tags: Aspose.Cells create pie chart from ListObject | batch add charts to Excel workbook .xlsx | set chart title using table display name Aspose.Cells | enable percentage data labels in Aspose.Cells pie chart | iterate worksheets and tables with Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace PieChartBatchGenerator
{
    // The program loads an existing .xlsx file, loops through every worksheet and each ListObject (Excel table), adds a pie chart positioned in rows 5‑20 and columns A‑K, sets the chart title to the table's display name, binds the chart series to the table's data range, enables percentage and category name data labels, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = @"C:\Data\SourceWorkbook.xlsx";
            string outputPath = @"C:\Data\WorkbookWithPieCharts.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process each ListObject (Excel table) on the worksheet
                    foreach (ListObject table in sheet.ListObjects)
                    {
                        // Define chart placement (rows, columns). Adjust as needed.
                        int upperLeftRow = 5;
                        int upperLeftColumn = 0;
                        int lowerRightRow = 20;
                        int lowerRightColumn = 10;

                        // Add a new pie chart to the worksheet (lifecycle rule: create)
                        int chartIndex = sheet.Charts.Add(ChartType.Pie, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
                        Chart chart = sheet.Charts[chartIndex];

                        // Use the table's display name for the chart title (fallback to a generic name if null)
                        string chartTitle = !string.IsNullOrEmpty(table.DisplayName) ? table.DisplayName : "Pie Chart";
                        chart.Title.Text = chartTitle;
                        chart.Title.IsVisible = true;

                        // Add the data series to the chart.
                        // The first column is assumed to be the category (X values) and the second column the values (Y values).
                        // Use the range address string required by NSeries.Add.
                        string dataRangeAddress = table.DataRange.RefersTo;
                        chart.NSeries.Add(dataRangeAddress, true);
                        chart.NSeries[0].Type = ChartType.Pie;

                        // Enable data labels to show percentages and category names
                        chart.NSeries[0].DataLabels.ShowPercentage = true;
                        chart.NSeries[0].DataLabels.ShowCategoryName = true;
                    }
                }

                // Ensure output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook with the new charts (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
