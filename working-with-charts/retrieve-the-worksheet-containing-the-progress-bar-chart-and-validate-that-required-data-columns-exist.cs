// Title: Find the worksheet that contains a 'Progress Bar' chart and validate required columns (Task, StartDate, EndDate) with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to iterate through all worksheets and return the name of the sheet that contains a chart titled 'Progress Bar'. | Create a method that loads an Excel workbook, locates the 'Progress Bar' chart, and checks whether the first row includes the headers Task, StartDate, and EndDate, returning a boolean indicating completeness. | Generate a reusable validator that logs any missing required columns and outputs the zero‑based index of the worksheet where the 'Progress Bar' chart is found.
// Common Searches: asp.net aspose.cells find worksheet by chart title progress bar | c# verify that Excel sheet has Task, StartDate, EndDate headers for a specific chart | how to locate chart named 'Progress Bar' in an Excel workbook using Aspose.Cells | validate required data columns for a progress bar chart in .NET Excel processing
// Tags: worksheet lookup by chart name Aspose.Cells | required column header validation Excel .NET | progress bar chart column check Aspose.Cells | chart presence verification C# | first row header extraction Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, scans each worksheet and its charts to find a chart titled 'Progress Bar', identifies the containing worksheet, reads the first row to collect header values, and confirms that the columns 'Task', 'StartDate', and 'EndDate' exist, reporting missing columns if any.
class ProgressBarChartValidator
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Locate the worksheet containing the "Progress Bar" chart
            Worksheet chartWorksheet = null;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Chart chart in sheet.Charts)
                {
                    if (string.Equals(chart.Name, "Progress Bar", StringComparison.OrdinalIgnoreCase))
                    {
                        chartWorksheet = sheet;
                        break;
                    }
                }
                if (chartWorksheet != null)
                    break;
            }

            if (chartWorksheet == null)
            {
                Console.WriteLine("The workbook does not contain a chart named 'Progress Bar'.");
                return;
            }

            Console.WriteLine($"Chart found in worksheet: {chartWorksheet.Name}");

            // Required column headers
            List<string> requiredHeaders = new List<string> { "Task", "StartDate", "EndDate" };

            // Assume the first row (index 0) contains headers
            Row headerRow = chartWorksheet.Cells.Rows[0];
            HashSet<string> existingHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (Cell cell in headerRow)
            {
                // Use Cell.Type to determine if the cell is empty
                if (cell.Type != CellValueType.IsNull)
                {
                    existingHeaders.Add(cell.StringValue.Trim());
                }
            }

            // Validate required columns
            bool allPresent = true;
            foreach (string header in requiredHeaders)
            {
                if (!existingHeaders.Contains(header))
                {
                    Console.WriteLine($"Missing required column: {header}");
                    allPresent = false;
                }
            }

            Console.WriteLine(allPresent
                ? "All required data columns are present."
                : "One or more required columns are missing.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
