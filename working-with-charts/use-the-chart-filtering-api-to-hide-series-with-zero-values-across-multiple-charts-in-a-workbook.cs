// Title: Hide Zero‑Value Chart Series Across All Charts in an Excel Workbook using Aspose.Cells for .NET
// Description: This example loads an Excel file, scans every worksheet and each chart, evaluates the data range of every series, and sets the series' IsFiltered flag to true when all numeric values are zero, effectively hiding empty series before saving the workbook.
// Keywords: Aspose.Cells | C# | Excel chart series filter | hide zero value series | IsFiltered property | multiple worksheets | chart automation | chart filtering API | remove empty series | Excel workbook processing
// Common Searches: Aspose.Cells hide chart series with zero values | filter out empty series from Excel charts C# | programmatically hide zero‑value series in multiple charts | use IsFiltered to hide chart series Aspose.Cells | iterate all charts in a workbook and hide zero data
// Developer Intent: Automatically hide any chart series whose numeric data consists solely of zeros across all charts in a workbook.
// Use Cases: Prepare financial reports by removing series that represent zero sales, keeping charts clean for stakeholders. | Generate presentation‑ready dashboards where only meaningful data series appear, improving visual clarity. | Automate Excel workbook cleanup in batch processes, ensuring charts do not display empty or placeholder series.
// AI Prompts: Create a reusable method that accepts a Workbook object and hides all chart series with only zero values using Aspose.Cells. | Explain the impact of the IsFiltered property on chart rendering and best practices for its use in .NET applications. | Suggest robust error‑handling enhancements for the chart‑filtering loop, including handling of non‑numeric cells, empty ranges, and missing charts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartFilteringDemo
{
    // This example loads an Excel file, scans every worksheet and each chart, evaluates the data range of every series, and sets the series' IsFiltered flag to true when all numeric values are zero, effectively hiding empty series before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook containing charts
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all charts on the worksheet
                    foreach (Chart chart in worksheet.Charts)
                    {
                        // Examine each series in the chart
                        for (int i = 0; i < chart.NSeries.Count; i++)
                        {
                            Series series = chart.NSeries[i];

                            // Get the data range of the series (e.g., "B2:B5")
                            string valuesRange = series.Values;

                            // Create a range object to access the cells
                            AsposeRange range = worksheet.Cells.CreateRange(valuesRange);

                            bool allZero = true;

                            // Check each cell in the range
                            foreach (Cell cell in range)
                            {
                                // Consider only numeric cells; ignore blanks or text
                                if (cell.Type == CellValueType.IsNumeric)
                                {
                                    if (cell.DoubleValue != 0)
                                    {
                                        allZero = false;
                                        break;
                                    }
                                }
                            }

                            // If all numeric values are zero, hide the series
                            if (allZero)
                            {
                                series.IsFiltered = true;
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
