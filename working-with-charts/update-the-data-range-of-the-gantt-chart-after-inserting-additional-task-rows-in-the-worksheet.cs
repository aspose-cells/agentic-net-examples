// Title: How to programmatically extend a Gantt chart data range after inserting task rows using Aspose.Cells for .NET
// AI Prompts: Insert rows at a specific index in an Excel worksheet and automatically adjust the Values formulas of all series in an existing Gantt chart to include the new rows with Aspose.Cells. | Calculate the worksheet's last data row after row insertion and rebuild each Gantt chart series range based on a given first data row using C# and Aspose.Cells. | Create a command‑line utility that accepts input file, sheet name, insert position, row count, chart index, and first data row, then updates the Gantt chart range and saves the workbook.
// Common Searches: Aspose.Cells C# update Gantt chart series after adding rows | programmatically change Excel chart data range when inserting rows using Aspose.Cells | C# example to expand Gantt chart range based on MaxDataRow | how to adjust chart series values formula after inserting rows in .NET | command line tool for updating Excel Gantt chart with Aspose.Cells
// Tags: Aspose.Cells Gantt chart series range update | refresh chart data after worksheet row addition .NET | dynamic chart range using MaxDataRow | command‑line Excel processing with Aspose.Cells | C# modify chart source after worksheet changes

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGanttUpdate
{
    // Loads an Excel workbook, inserts task rows at a specified index, determines the new last data row, updates each series' Values formula in the targeted Gantt chart to cover the expanded range, and saves the modified file—all configurable via command‑line arguments.
    class Program
    {
        static void Main(string[] args)
        {
            // Input parameters – can be passed via command line or use defaults
            string inputFilePath = args.Length > 0 ? args[0] : "input.xlsx";
            string outputFilePath = args.Length > 1 ? args[1] : "output.xlsx";
            string sheetName = args.Length > 2 ? args[2] : "Sheet1";
            int insertRowIndex = args.Length > 3 ? int.Parse(args[3]) : 1;      // zero‑based index
            int rowsToInsert = args.Length > 4 ? int.Parse(args[4]) : 1;
            int chartIndex = args.Length > 5 ? int.Parse(args[5]) : 0;
            int firstDataRow = args.Length > 6 ? int.Parse(args[6]) : 2;       // first row of data (1‑based)

            // Verify input file exists
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Input file not found: {inputFilePath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputFilePath);

                // Get the worksheet that contains the Gantt chart
                Worksheet worksheet = workbook.Worksheets[sheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                    return;
                }

                // Insert new task rows
                worksheet.Cells.InsertRows(insertRowIndex, rowsToInsert);

                // Validate chart index
                if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                {
                    Console.WriteLine($"Chart index {chartIndex} is out of range.");
                    return;
                }

                // Get the Gantt chart
                Chart ganttChart = worksheet.Charts[chartIndex];
                if (ganttChart == null)
                {
                    Console.WriteLine("Specified chart could not be retrieved.");
                    return;
                }

                // Determine the new data range boundaries
                int lastDataRow = worksheet.Cells.MaxDataRow;

                // Update each series in the chart to reference the expanded range
                foreach (Series series in ganttChart.NSeries)
                {
                    try
                    {
                        // Use Values formula for the series (Gantt charts are typically bar charts)
                        string formula = series.Values; // e.g., Sheet1!$B$2:$B$5
                        if (string.IsNullOrEmpty(formula))
                            continue;

                        // Extract column letter(s) from the first cell reference
                        int exclPos = formula.IndexOf('!');
                        if (exclPos < 0)
                            continue; // unexpected format

                        string rangePart = formula.Substring(exclPos + 1); // "$B$2:$B$5"
                        string[] parts = rangePart.Split(':');
                        if (parts.Length == 0)
                            continue;

                        string startRef = parts[0]; // "$B$2"
                        string columnLetter = new string(startRef.Where(Char.IsLetter).ToArray());

                        // Build the new formula with updated row numbers (no leading '=')
                        string newFormula = $"{worksheet.Name}!${columnLetter}${firstDataRow}:${columnLetter}${lastDataRow}";
                        series.Values = newFormula;
                    }
                    catch (Exception exSeries)
                    {
                        Console.WriteLine($"Failed to update a series: {exSeries.Message}");
                    }
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook with the updated chart
                workbook.Save(outputFilePath);
                Console.WriteLine($"Workbook saved successfully to {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
