// Title: Refresh a Pivot Table After Expanding Its Named Range with Aspose.Cells (C#)
// Description: This example demonstrates how to load a workbook, append new rows to the source sheet, enlarge the "DataRange" named range, refresh all pivot tables on a target sheet, and save the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# refresh pivot table | expand named range Excel | RefreshPivotTables method | update DataRange programmatically | dynamic pivot data source | Aspose.Cells named range example | C# Excel pivot table automation | add rows to worksheet Aspose.Cells | pivot table source range change
// Common Searches: how to refresh pivot table after adding rows Aspose.Cells | expand named range and refresh pivot C# | programmatically change pivot source range Aspose.Cells | Aspose.Cells RefreshPivotTables usage | update Excel named range with Aspose.Cells
// Developer Intent: Programmatically extend the data range used by a pivot table and trigger a refresh so the new rows are reflected in the pivot report.
// Use Cases: Add daily sales entries, grow the "DataRange" named range, and automatically refresh the sales‑summary pivot on the dashboard sheet. | Import a new month of financial figures, adjust the source named range, and refresh the monthly financial pivot without manual intervention. | Insert additional product records into a catalog, update the associated named range, and ensure all analysis pivots display the latest data.
// AI Prompts: Generate C# code that appends rows to a worksheet, expands a named range, and calls RefreshPivotTables on a specific sheet using Aspose.Cells. | Show an Aspose.Cells example that checks for a named range, creates it if missing, updates its RefersTo property after data insertion, and refreshes all pivot tables. | Write a reusable function that receives a workbook path, a list of new rows, and a named range name, then adds the rows, expands the range, refreshes pivots, and saves the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example demonstrates how to load a workbook, append new rows to the source sheet, enlarge the "DataRange" named range, refresh all pivot tables on a target sheet, and save the updated file using Aspose.Cells for .NET.
class RefreshPivotAfterRangeExpand
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

            // Load the existing workbook that contains the data range and the pivot table
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // 1. Add new rows to the data source (first sheet)
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0]; // assume data is on the first sheet
            int newRowIndex = dataSheet.Cells.MaxDataRow + 1; // first empty row after existing data

            // Example new data – adjust columns as needed
            dataSheet.Cells[newRowIndex, 0].PutValue("NewProduct"); // Column A
            dataSheet.Cells[newRowIndex, 1].PutValue(250);         // Column B

            // -------------------------------------------------
            // 2. Expand the named range that the pivot table uses
            // -------------------------------------------------
            // Assume the named range is called "DataRange" and originally starts at A1
            Name dataRange = workbook.Worksheets.Names["DataRange"];

            // If the named range does not exist, create it
            if (dataRange == null)
            {
                dataRange = workbook.Worksheets.Names[workbook.Worksheets.Names.Add("DataRange")];
            }

            // Build the new address string: =SheetName!$A$1:$B${lastRow}
            // Column B has index 1 (zero‑based). CellsHelper converts indexes to Excel style names.
            string lastCellAddress = CellsHelper.CellIndexToName(1, newRowIndex); // e.g., "B5"
            string newRefersTo = $"={dataSheet.Name}!$A$1:${lastCellAddress}";

            dataRange.RefersTo = newRefersTo; // update the named range to include the new rows

            // -------------------------------------------------
            // 3. Refresh the pivot table so it picks up the expanded range
            // -------------------------------------------------
            // Assume the pivot table resides on the second worksheet
            if (workbook.Worksheets.Count > 1)
            {
                Worksheet pivotSheet = workbook.Worksheets[1];
                pivotSheet.RefreshPivotTables(); // refreshes all pivot tables on this sheet
            }
            else
            {
                Console.WriteLine("Pivot sheet not found.");
            }

            // -------------------------------------------------
            // 4. Save the updated workbook
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
