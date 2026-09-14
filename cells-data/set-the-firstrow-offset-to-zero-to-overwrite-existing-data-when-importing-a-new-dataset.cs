// Title: Overwrite existing worksheet rows by importing a DataTable with firstRow offset set to zero using Aspose.Cells for .NET
// AI Prompts: Import a DataTable into a worksheet at row 0 and replace any existing cells using ImportTableOptions. | Configure ImportTableOptions with ShiftFirstRowDown = false and InsertRows = false to ensure overwriting when calling Cells.ImportData. | Save the workbook after the second import so the Excel file contains only the new dataset.
// Common Searches: how to overwrite Excel sheet data with a new DataTable using Aspose.Cells C# | Aspose.Cells ImportData starting at first row zero overwrite existing rows | set ShiftFirstRowDown false to prevent row shifting during import in Aspose.Cells | reuse ImportTableOptions to replace worksheet content in .NET | C# import DataTable into existing workbook without inserting rows
// Tags: Aspose.Cells ImportData overwrite rows | ImportTableOptions ShiftFirstRowDown false | InsertRows false Aspose.Cells | DataTable import starting at row zero | replace worksheet content C# Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsImportOverwriteDemo
{
    // The example creates a workbook, imports an initial DataTable at row 0, then imports a second DataTable using the same ImportTableOptions (ShiftFirstRowDown = false, InsertRows = false) to overwrite the previous content, and finally saves the result to OverwriteImportResult.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // -----------------------------------------------------------------
            // First import – populate the sheet with initial data (row offset 0)
            // -----------------------------------------------------------------
            DataTable initialTable = new DataTable();
            initialTable.Columns.Add("Product");
            initialTable.Columns.Add("Quantity", typeof(int));
            initialTable.Rows.Add("Apple", 10);
            initialTable.Rows.Add("Banana", 20);

            // Default import options (ShiftFirstRowDown = false, InsertRows = false)
            ImportTableOptions initialOptions = new ImportTableOptions
            {
                IsFieldNameShown = false,   // Do not import column names
                ShiftFirstRowDown = false,  // Do not shift the first row down
                InsertRows = false          // Overwrite existing cells
            };

            // Import the initial data starting at row 0, column 0
            cells.ImportData(initialTable, 0, 0, initialOptions);

            // ---------------------------------------------------------------
            // Second import – new dataset should overwrite the previous content
            // ---------------------------------------------------------------
            DataTable newTable = new DataTable();
            newTable.Columns.Add("Product");
            newTable.Columns.Add("Quantity", typeof(int));
            newTable.Rows.Add("Orange", 15);
            newTable.Rows.Add("Grape", 25);
            newTable.Rows.Add("Mango", 30);

            // Reuse the same options to ensure overwriting (firstRow offset = 0)
            ImportTableOptions overwriteOptions = new ImportTableOptions
            {
                IsFieldNameShown = false,
                ShiftFirstRowDown = false,
                InsertRows = false
            };

            // Import the new data starting again at row 0, column 0
            cells.ImportData(newTable, 0, 0, overwriteOptions);

            // Save the workbook – the sheet now contains only the new dataset
            workbook.Save("OverwriteImportResult.xlsx");
        }
    }
}
