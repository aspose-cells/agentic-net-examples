// Title: How to update a Gantt chart’s data range after inserting task rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# expand Gantt chart series after inserting rows | update chart data source range after adding tasks in Excel workbook with Aspose.Cells | C# code to adjust Gantt chart range when new task rows are inserted | modify chart series formula dynamically in Aspose.Cells | insert rows and refresh Gantt chart in Aspose.Cells .NET
// Tags: chart series range expansion Aspose.Cells | row insertion impact on Gantt chart Aspose.Cells | modify chart data source C# Aspose.Cells | dynamic Gantt chart data range .NET | update chart formulas after inserting rows

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads a workbook containing a Gantt chart, inserts new task rows at a specified position, expands each chart series' data range formula to include the inserted rows, and saves the workbook with the updated chart.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "GanttChart.xlsx";
            const string outputFile = "GanttChart_Updated.xlsx";

            // Verify that the template workbook exists
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file '{inputFile}' was not found.");

            // Load the existing workbook that contains the Gantt chart
            Workbook workbook = new Workbook(inputFile);
            Worksheet sheet = workbook.Worksheets[0];

            // Define where new task rows will be inserted and how many rows
            int insertRowIndex = 5;   // zero‑based index where rows are inserted
            int rowsToInsert = 2;     // number of new task rows

            // Insert the rows into the worksheet
            sheet.Cells.InsertRows(insertRowIndex, rowsToInsert);

            // (Optional) Populate the newly inserted rows with task data
            // Example:
            // sheet.Cells[insertRowIndex, 0].PutValue("New Task 1");
            // sheet.Cells[insertRowIndex, 1].PutValue(DateTime.Today);
            // sheet.Cells[insertRowIndex, 2].PutValue(3); // duration

            // Retrieve the Gantt chart (assumed to be the first chart on the sheet)
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found on the worksheet.");

            Chart ganttChart = sheet.Charts[0];

            // Update each series in the chart to extend its data range to include the new rows
            foreach (Series series in ganttChart.NSeries)
            {
                // Current formula for Y values, e.g. =Sheet1!$B$2:$B$10
                string currentFormula = series.Values;

                // Expand the range to cover the inserted rows
                string updatedFormula = ExpandRange(currentFormula, insertRowIndex, rowsToInsert);

                // Apply the new formula to the series
                series.Values = updatedFormula;
            }

            // Save the workbook with the updated chart
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method that expands a range in a chart series formula to include newly inserted rows.
    // This simple implementation assumes the range is a single contiguous block like $B$2:$B$10.
    static string ExpandRange(string formula, int insertRowIndex, int rowsInserted)
    {
        if (string.IsNullOrEmpty(formula))
            return formula;

        // Find the colon that separates start and end addresses
        int colonPos = formula.IndexOf(':');
        if (colonPos < 0)
            return formula; // Not a range; nothing to expand

        // Split into start and end parts
        string startPart = formula.Substring(0, colonPos);
        string endPart = formula.Substring(colonPos + 1);

        // Locate the row number in the end part (after the last '$')
        int lastDollar = endPart.LastIndexOf('$');
        if (lastDollar < 0)
            return formula; // Unexpected format

        // Extract the numeric row component
        string rowNumberStr = endPart.Substring(lastDollar + 1);
        if (!int.TryParse(rowNumberStr, out int endRow))
            return formula; // Unable to parse row number

        // If the original end row is at or below the insertion point, shift it down
        if (endRow >= insertRowIndex + 1) // +1 because Excel rows are 1‑based
        {
            endRow += rowsInserted;
            // Rebuild the end address with the new row number
            string columnPart = endPart.Substring(0, lastDollar + 1);
            endPart = columnPart + endRow.ToString();
        }

        // Reconstruct the full formula
        return startPart + ":" + endPart;
    }
}
