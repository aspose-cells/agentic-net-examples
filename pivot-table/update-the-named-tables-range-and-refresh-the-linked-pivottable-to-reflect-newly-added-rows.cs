// Title: Resize a Named ListObject Table and Refresh Linked PivotTables with Aspose.Cells for .NET
// Description: Load a workbook, locate the ListObject named "SalesData", expand its range to rows 1‑20 and columns A‑C, then refresh every PivotTable so the new rows are included. The example handles missing files or tables, recalculates each pivot's range and data, and saves the updated workbook.
// Keywords: Aspose.Cells C# resize ListObject | update named table range | refresh PivotTables programmatically | Excel table expansion Aspose | pivot cache refresh .NET | Workbook automation Aspose.Cells | ListObject Resize method | PivotTable CalculateRange | PivotTable CalculateData | Excel data source change
// Common Searches: How to change the range of a ListObject and refresh its pivot tables using Aspose.Cells | Aspose.Cells C# resize named table and recalculate pivots | Update Excel table source range and refresh all pivots programmatically | Resize ListObject and call RefreshPivotTables in .NET | Expand sales data table and update linked pivot chart with Aspose
// Developer Intent: Expand the "SalesData" ListObject to a larger area and automatically refresh all associated PivotTables so they reflect the added rows.
// Use Cases: Monthly data import adds new rows; the table must grow and the sales summary pivot must update before distribution. | Dynamic reporting where the number of records varies per client workbook; automate table resizing and pivot refresh. | Batch processing of multiple workbooks to adjust table ranges after data cleansing and ensure accurate pivot calculations.
// AI Prompts: Write C# code with Aspose.Cells that resizes a ListObject called 'SalesData' to include rows 1‑30 and refreshes every PivotTable in the workbook. | Show an example that expands a named table, calls RefreshPivotTables, then invokes CalculateRange and CalculateData on each pivot using Aspose.Cells for .NET. | Explain how to add robust error handling when the specified ListObject does not exist before resizing and refreshing pivots.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Load a workbook, locate the ListObject named "SalesData", expand its range to rows 1‑20 and columns A‑C, then refresh every PivotTable so the new rows are included. The example handles missing files or tables, recalculates each pivot's range and data, and saves the updated workbook.
    public class UpdateTableAndRefreshPivot
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file '{inputPath}' not found.");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the table is on the first worksheet and named "SalesData"
            Worksheet dataSheet = workbook.Worksheets[0];
            const string tableName = "SalesData";

            // Retrieve the ListObject (named table) by its name
            ListObject table = dataSheet.ListObjects[tableName];
            if (table == null)
                throw new InvalidOperationException($"Table '{tableName}' not found on worksheet '{dataSheet.Name}'.");

            // Define the new data range for the table (e.g., expand to rows 1-20, columns A-C)
            // Cell indices are zero‑based: A1 -> (0,0), C20 -> (19,2)
            CellArea newRange = new CellArea
            {
                StartRow = 0,
                StartColumn = 0,
                EndRow = 19,
                EndColumn = 2
            };

            // Resize the table using start row/column, total rows/columns, and indicate that the table has headers
            int totalRows = newRange.EndRow - newRange.StartRow + 1;      // 20 rows
            int totalColumns = newRange.EndColumn - newRange.StartColumn + 1; // 3 columns
            table.Resize(newRange.StartRow, newRange.StartColumn, totalRows, totalColumns, true);

            // Refresh all pivot tables in the workbook
            workbook.Worksheets.RefreshPivotTables();

            // Ensure each pivot table recalculates its source range and data
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    pt.CalculateRange();   // Recalculate the pivot table's range
                    pt.CalculateData();    // Recalculate the pivot results
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
