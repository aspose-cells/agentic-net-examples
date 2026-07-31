// Title: Refresh an Excel Pivot Table After Expanding Its Named Range with Aspose.Cells (C#)
// Description: Shows how to load a workbook, append rows to a data sheet, enlarge the "SourceData" named range, refresh pivot tables on another sheet, and save the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel pivot table refresh | named range expansion | update pivot source | RefreshPivotTables | Workbook.Save | data worksheet | pivot worksheet | programmatic Excel automation
// Common Searches: Aspose.Cells refresh pivot after adding rows | C# expand named range for pivot table | how to update pivot source range in Aspose.Cells | refresh all pivot tables in a workbook using Aspose.Cells | add rows to Excel sheet and refresh pivot programmatically
// Developer Intent: Programmatically extend the source named range and trigger a refresh so the pivot tables include the newly added rows.
// Use Cases: Automated daily data loads where new records are appended and reporting pivots must reflect them instantly. | Batch processing of Excel files that require dynamic pivot updates after importing additional data. | Server‑side .NET services that generate refreshed pivot‑based dashboards on demand.
// AI Prompts: Generate C# code with Aspose.Cells that adds rows to a worksheet, expands a named range, refreshes pivot tables on another sheet, and saves the workbook. | Explain step‑by‑step how to modify a named range and refresh associated pivot tables in an Aspose.Cells workbook after inserting new data rows. | Provide a concise example that loads an existing Excel file, updates the source data, adjusts the "SourceData" reference, calls RefreshPivotTables, and writes the result to disk.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to load a workbook, append rows to a data sheet, enlarge the "SourceData" named range, refresh pivot tables on another sheet, and save the updated file using Aspose.Cells for .NET.
class RefreshPivotAfterRangeExpand
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists.
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook that contains the source data and the pivot table.
            Workbook workbook = new Workbook(inputPath);

            // -------------------------------------------------
            // 1. Append new rows to the source data worksheet.
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets["Data"]; // assume the data sheet is named "Data"
            if (dataSheet == null)
                throw new InvalidOperationException("Worksheet named 'Data' was not found.");

            // Determine the last row that currently contains data (zero‑based index).
            int lastDataRow = dataSheet.Cells.MaxDataRow;

            // Add two new rows of sample data.
            dataSheet.Cells[lastDataRow + 1, 0].PutValue("NewItem1"); // Column A
            dataSheet.Cells[lastDataRow + 1, 1].PutValue(123);       // Column B

            dataSheet.Cells[lastDataRow + 2, 0].PutValue("NewItem2"); // Column A
            dataSheet.Cells[lastDataRow + 2, 1].PutValue(456);       // Column B

            // -------------------------------------------------
            // 2. Expand the named range used by the pivot table.
            // -------------------------------------------------
            // Assume the named range that the pivot table references is called "SourceData".
            Name sourceRange = workbook.Worksheets.Names["SourceData"];
            if (sourceRange == null)
                throw new InvalidOperationException("Named range 'SourceData' was not found.");

            // Calculate the new address for the expanded range.
            // Columns A (0) to B (1) are used; rows start at 1 in Excel notation.
            int newLastRowIndex = lastDataRow + 2;               // zero‑based index of the new last row
            int newLastRowNumber = newLastRowIndex + 1;          // Excel row number (1‑based)
            string newAddress = $"=Data!$A$1:${CellsHelper.ColumnIndexToName(1)}${newLastRowNumber}";

            // Update the named range to point to the expanded area.
            sourceRange.RefersTo = newAddress;

            // -------------------------------------------------
            // 3. Refresh the pivot table so it picks up the new rows.
            // -------------------------------------------------
            // Assume the pivot table resides on a sheet named "Pivot".
            Worksheet pivotSheet = workbook.Worksheets["Pivot"];
            if (pivotSheet == null)
                throw new InvalidOperationException("Worksheet named 'Pivot' was not found.");

            pivotSheet.RefreshPivotTables(); // Refreshes all pivot tables on this worksheet.

            // -------------------------------------------------
            // 4. Save the workbook with the refreshed pivot table.
            // -------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
