// Title: Add a line chart to every worksheet and link its data labels to source cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add line chart to each worksheet in a workbook | How to bind chart data labels to the original cells with Aspose.Cells .NET | Programmatically set line chart series and category ranges for all sheets using Aspose.Cells | Batch create line charts that show values and categories as data labels in Excel with Aspose.Cells | Iterate through worksheets and add charts with data labels linked to cell ranges in Aspose.Cells
// Tags: line chart creation per worksheet Aspose.Cells | data label cell linking Aspose.Cells | chart series range assignment C# | batch chart generation Excel Aspose.Cells | category and value range definition Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace BatchLineChartDemo
{
    // The example loads an existing Excel workbook, loops through each worksheet, determines the used data range, adds a line chart positioned between rows 5‑20 and columns A‑J, assigns column B as the series values and column A as categories, enables data labels to display both the numeric value and the category name linked directly to the source cells, centers the labels, sets a sheet‑specific title, and saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range of the worksheet
                    int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
                    int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last column with data

                    // We expect at least two columns: categories (A) and values (B)
                    if (maxRow < 1 || maxCol < 1)
                    {
                        Console.WriteLine($"Worksheet \"{sheet.Name}\" does not contain enough data for a chart. Skipping.");
                        continue;
                    }

                    // Build the address strings for the series values and categories.
                    // Assuming the first row (row 0) contains headers.
                    // Values are in column B (index 1), categories in column A (index 0).
                    string valueRange = $"B2:B{maxRow + 1}";
                    string categoryRange = $"A2:A{maxRow + 1}";

                    // Add a line chart to the worksheet.
                    // Parameters: ChartType, upper left row, upper left column, lower right row, lower right column
                    int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                    Chart chart = sheet.Charts[chartIndex];

                    // Set the data source for the chart.
                    chart.NSeries.Add(valueRange, true);
                    chart.NSeries.CategoryData = categoryRange;

                    // Access the first (and only) series.
                    Series series = chart.NSeries[0];

                    // Enable data labels and link them to the source cells.
                    series.DataLabels.ShowValue = true;            // Show the numeric value from the cell.
                    series.DataLabels.ShowCategoryName = true;    // Show the category (e.g., A‑column label).
                    series.DataLabels.ShowSeriesName = false;     // Not needed for a single series.

                    // Optional: position the data labels.
                    series.DataLabels.Position = LabelPositionType.Center;

                    // Give the chart a meaningful title.
                    chart.Title.Text = $"Line Chart – {sheet.Name}";
                    chart.Title.IsVisible = true;

                    Console.WriteLine($"Added line chart to worksheet \"{sheet.Name}\" using range {valueRange}.");
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
