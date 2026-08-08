// Title: Import a DataTable Overwrite Existing Cells at Row 0 with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to import a DataTable into a worksheet starting at the first row (offset 0) and replace existing content. The example configures ImportTableOptions (ShiftFirstRowDown = false, InsertRows = false, IsFieldNameShown = false) so the new data overwrites cells A1‑A3 without shifting rows or adding headers, then saves the workbook.
// Keywords: Aspose.Cells import DataTable C# | overwrite existing cells Aspose.Cells | ImportData row offset zero | ShiftFirstRowDown false | InsertRows false | ImportTableOptions overwrite | Aspose.Cells .NET example
// Common Searches: Aspose.Cells import DataTable without shifting rows | How to overwrite cells when using ImportData in Aspose.Cells | ImportData starting at first row C# | Prevent row insertion on ImportData Aspose.Cells | Replace worksheet data with new DataTable Aspose
// Developer Intent: Replace the current worksheet content by importing a new DataTable at the first row, overwriting existing cells.
// Use Cases: Refresh a report template with new data while preserving layout. | Update a dashboard sheet by overwriting previous calculations with fresh results. | Replace old headers and rows in a generated spreadsheet without adding extra rows.
// AI Prompts: Show C# code to import a DataTable into an Aspose.Cells worksheet at row 0 and overwrite existing cells. | Explain how to set ImportTableOptions to prevent row shifting and insertion during ImportData. | Provide guidance on using ShiftFirstRowDown and InsertRows properties for an overwrite import in Aspose.Cells.

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportOverwriteDemo
{
    // Demonstrates how to import a DataTable into a worksheet starting at the first row (offset 0) and replace existing content. The example configures ImportTableOptions (ShiftFirstRowDown = false, InsertRows = false, IsFieldNameShown = false) so the new data overwrites cells A1‑A3 without shifting rows or adding headers, then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();                     // create workbook
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // Step 1: Add some existing data that we will later overwrite
            // ------------------------------------------------------------
            cells["A1"].PutValue("Old Header");
            cells["A2"].PutValue("Old Data 1");
            cells["A3"].PutValue("Old Data 2");

            // ------------------------------------------------------------
            // Step 2: Prepare a new DataTable to import
            // ------------------------------------------------------------
            DataTable newTable = new DataTable();
            newTable.Columns.Add("Header", typeof(string));
            newTable.Columns.Add("Value", typeof(string));

            // New data rows (these will replace the old ones)
            newTable.Rows.Add("New Header", "New Data 1");
            newTable.Rows.Add("New Header", "New Data 2");

            // ------------------------------------------------------------
            // Step 3: Configure import options to overwrite existing cells
            // ------------------------------------------------------------
            ImportTableOptions importOptions = new ImportTableOptions
            {
                // Do not shift the first row down; start writing at the exact row we specify
                ShiftFirstRowDown = false,
                // Do not insert new rows; write over existing rows
                InsertRows = false,
                // Do not write field names (headers) because they are already part of the data rows
                IsFieldNameShown = false
            };

            // ------------------------------------------------------------
            // Step 4: Import the new data starting at row 0, column 0 (A1)
            // ------------------------------------------------------------
            // Because ShiftFirstRowDown = false and InsertRows = false,
            // the import will overwrite the cells that already contain "Old Header", etc.
            cells.ImportData(newTable, 0, 0, importOptions);

            // ------------------------------------------------------------
            // Step 5: Save the workbook to verify the result
            // ------------------------------------------------------------
            workbook.Save("OverwriteImportDemo.xlsx"); // save workbook
        }
    }
}
