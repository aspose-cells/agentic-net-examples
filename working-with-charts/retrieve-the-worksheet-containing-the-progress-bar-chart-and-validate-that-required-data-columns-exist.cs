// Title: Locate the worksheet that contains a 'Progress Bar' chart and verify required data columns with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells find worksheet that has a chart named Progress Bar | How to search all worksheets for a specific chart title using Aspose.Cells in .NET | Validate that Task, StartDate, EndDate, and Progress columns exist in the sheet backing a progress bar chart | Retrieve worksheet containing a progress bar chart and ensure required data columns are present with Aspose.Cells | Aspose.Cells code to check required column headers in Excel sheet that holds a progress bar chart
// Tags: retrieve worksheet by chart title Aspose.Cells | filter charts by name Aspose.Cells .NET | validate required column headers Aspose.Cells | detect progress bar chart in Excel using C# | check data column existence Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, iterates through each worksheet to locate a chart whose title or name includes "Progress Bar", captures that worksheet, and then scans the first row for the headers "Task", "StartDate", "EndDate", and "Progress", reporting any missing columns.
class ProgressBarChartValidator
{
    static void Main()
    {
        try
        {
            // Path to the workbook that contains the Progress Bar chart
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' does not exist.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Variable to hold the worksheet that contains the Progress Bar chart
            Worksheet progressBarWorksheet = null;

            // Search all worksheets for a chart whose name contains "Progress Bar"
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Each worksheet may contain multiple charts
                foreach (Chart chart in sheet.Charts)
                {
                    // Chart name can be retrieved via the Title property or Name property
                    string chartName = chart.Title?.Text ?? chart.Name;

                    if (!string.IsNullOrEmpty(chartName) &&
                        chartName.IndexOf("Progress Bar", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        progressBarWorksheet = sheet;
                        break;
                    }
                }

                if (progressBarWorksheet != null)
                    break;
            }

            // If the worksheet was not found, report and exit
            if (progressBarWorksheet == null)
            {
                Console.WriteLine("Error: No worksheet containing a 'Progress Bar' chart was found.");
                return;
            }

            Console.WriteLine($"Progress Bar chart found in worksheet: '{progressBarWorksheet.Name}'");

            // Define the required data column headers (adjust as needed)
            string[] requiredColumns = { "Task", "StartDate", "EndDate", "Progress" };

            // Assume headers are located in the first row (row index 0)
            int headerRowIndex = 0;

            // Collect the column indices of the required headers
            Dictionary<string, int> columnIndices = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // Scan the first row for each required header
            foreach (string header in requiredColumns)
            {
                bool found = false;
                int maxColumn = progressBarWorksheet.Cells.MaxColumn;

                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = progressBarWorksheet.Cells[headerRowIndex, col];
                    if (cell != null && cell.Type == CellValueType.IsString &&
                        string.Equals(cell.StringValue.Trim(), header, StringComparison.OrdinalIgnoreCase))
                    {
                        columnIndices[header] = col;
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    Console.WriteLine($"Error: Required column '{header}' is missing in worksheet '{progressBarWorksheet.Name}'.");
                }
            }

            // Report validation result
            if (columnIndices.Count == requiredColumns.Length)
            {
                Console.WriteLine("All required data columns are present.");
            }
            else
            {
                Console.WriteLine("One or more required columns are missing. Validation failed.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
