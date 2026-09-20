// Title: Extend Gantt chart series ranges after inserting task rows using Aspose.Cells for .NET (C#)
// AI Prompts: Programmatically expand each series' Values range in an existing Gantt chart after adding new rows with Aspose.Cells in C#. | Recalculate and set the chart formula strings to include inserted rows when updating an Excel workbook via Aspose.Cells.
// Common Searches: Aspose.Cells C# update chart series range after inserting rows in worksheet | how to extend Gantt chart data range dynamically in .NET Excel file | adjust Excel chart formulas when adding task rows using Aspose.Cells | C# code to modify chart Values formula after row insertion | increase Gantt chart series end row in Aspose.Cells workbook
// Tags: chart series range extension Aspose.Cells | gantt chart data range update C# | row insertion impact on Excel chart Aspose.Cells | dynamic chart range handling .NET | update chart formulas programmatically

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GanttChartUpdater
{
    // The example loads a workbook, inserts additional task rows on the 'Tasks' sheet, iterates through each series of the first chart, parses the original Values formula, expands the end row by the number of inserted rows, rebuilds the absolute range string, assigns the new formula to the series, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "GanttChart.xlsx";
                const string outputPath = "GanttChart_Updated.xlsx";

                // Verify that the source workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Get the worksheet that contains the Gantt chart data
                var worksheet = workbook.Worksheets["Tasks"]; // adjust sheet name as needed
                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet 'Tasks' not found.");
                    return;
                }

                // Insert additional task rows (e.g., insert 3 rows after row 5)
                int insertAfterRowIndex = 5;          // zero‑based index of the row after which new rows are added
                int rowsToInsert = 3;
                worksheet.Cells.InsertRows(insertAfterRowIndex + 1, rowsToInsert);

                // Retrieve the Gantt chart (assumed to be the first chart on the sheet)
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the worksheet.");
                    return;
                }

                var chart = worksheet.Charts[0];

                // Update each series in the chart to extend its data range to include the newly inserted rows
                foreach (var series in chart.NSeries)
                {
                    try
                    {
                        // Example formula: "=Tasks!$B$2:$B$10"
                        string originalFormula = series.Values; // use Values property for the data range

                        // Find the '!' separating sheet name and range
                        int exclPos = originalFormula.IndexOf('!');
                        if (exclPos < 0)
                            continue; // unexpected format

                        // Extract the range part (e.g., "$B$2:$B$10")
                        string rangePart = originalFormula.Substring(exclPos + 1).Trim('\'');

                        // Create a range object from the string
                        var range = worksheet.Cells.CreateRange(rangePart);

                        // Determine start and end coordinates of the original range
                        int startRow = range.FirstRow;
                        int startColumn = range.FirstColumn;
                        int endRow = startRow + range.RowCount - 1;
                        int endColumn = startColumn + range.ColumnCount - 1;

                        // Extend the end row by the number of inserted rows
                        endRow += rowsToInsert;

                        // Build absolute cell names (e.g., $B$2)
                        string startAbs = $"${CellsHelper.ColumnIndexToName(startColumn)}${startRow + 1}";
                        string endAbs = $"${CellsHelper.ColumnIndexToName(endColumn)}${endRow + 1}";
                        string newRange = $"{startAbs}:{endAbs}";

                        // Update the series values formula
                        series.Values = $"={worksheet.Name}!{newRange}";
                    }
                    catch (Exception innerEx)
                    {
                        Console.WriteLine($"Failed to update a series: {innerEx.Message}");
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
