// Title: C# – Export Visible Rows and Columns to JSON with Aspose.Cells JsonSaveOptions
// Description: Demonstrates how to hide rows and columns in a workbook, copy only the visible cells to a temporary worksheet, configure JsonSaveOptions (ExportArea and SheetIndexes) and save the result as a JSON file that excludes all hidden rows and columns.
// Keywords: Aspose.Cells JsonSaveOptions C# | export visible rows to JSON | exclude hidden columns JSON | Excel to JSON without hidden data | C# Aspose.Cells export area | skip hidden rows Aspose.Cells
// Common Searches: Aspose.Cells export only visible cells to JSON | C# JsonSaveOptions hide hidden rows and columns | How to ignore hidden rows when saving JSON with Aspose.Cells | Export Excel worksheet to JSON excluding hidden columns | Aspose.Cells JsonSaveOptions ExportArea example
// Developer Intent: Generate a JSON file from an Excel workbook that contains only the rows and columns visible to the user.
// Use Cases: Create API responses that must not expose internal or hidden spreadsheet data. | Provide front‑end grids with clean JSON payloads that respect user‑hidden rows/columns. | Produce lightweight JSON reports where hidden rows are used for calculations only.
// AI Prompts: Show a C# example that exports only visible cells to JSON using Aspose.Cells without a temporary worksheet. | Explain how to set ExportArea and SheetIndexes in JsonSaveOptions to skip hidden rows and columns. | Give step‑by‑step code for exporting an Excel sheet to JSON while ignoring hidden rows/columns in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Demonstrates how to hide rows and columns in a workbook, copy only the visible cells to a temporary worksheet, configure JsonSaveOptions (ExportArea and SheetIndexes) and save the result as a JSON file that excludes all hidden rows and columns.
class JsonExportHiddenRowsColumns
{
    static void Main()
    {
        // Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Header row
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["C1"].PutValue("Header3");

        // Data rows
        sheet.Cells["A2"].PutValue("R1C1");
        sheet.Cells["B2"].PutValue("R1C2");
        sheet.Cells["C2"].PutValue("R1C3");
        sheet.Cells["A3"].PutValue("R2C1");
        sheet.Cells["B3"].PutValue("R2C2");
        sheet.Cells["C3"].PutValue("R2C3");

        // Hide a row (index 1 -> second row) and a column (index 1 -> column B)
        sheet.Cells.HideRow(1);
        sheet.Cells.HideColumn(1);

        // -----------------------------------------------------------------
        // Prepare JsonSaveOptions to export only visible rows and columns
        // -----------------------------------------------------------------
        JsonSaveOptions jsonOptions = new JsonSaveOptions();

        // Create a temporary worksheet that contains only the visible cells
        Worksheet tempSheet = workbook.Worksheets.Add("TempVisible");
        int destRow = 0;

        for (int r = 0; r <= sheet.Cells.MaxDataRow; r++)
        {
            // Skip hidden rows
            if (sheet.Cells.IsRowHidden(r))
                continue;

            int destCol = 0;
            for (int c = 0; c <= sheet.Cells.MaxDataColumn; c++)
            {
                // Skip hidden columns
                if (sheet.Cells.IsColumnHidden(c))
                    continue;

                // Copy the cell value to the temporary sheet
                object val = sheet.Cells[r, c].Value;
                tempSheet.Cells[destRow, destCol].PutValue(val);
                destCol++;
            }
            destRow++;
        }

        // Define the export area to cover the used range of the temporary sheet
        jsonOptions.ExportArea = new CellArea
        {
            StartRow = 0,
            StartColumn = 0,
            EndRow = tempSheet.Cells.MaxDataRow,
            EndColumn = tempSheet.Cells.MaxDataColumn
        };

        // Export only the temporary sheet (its index is the last one in the collection)
        jsonOptions.SheetIndexes = new int[] { workbook.Worksheets.Count - 1 };

        // Save the workbook as JSON using the configured options
        workbook.Save("ExportedVisibleData.json", jsonOptions);

        // Optional: clean up the temporary sheet if further processing is needed
        workbook.Worksheets.RemoveAt(workbook.Worksheets.Count - 1);
    }
}
