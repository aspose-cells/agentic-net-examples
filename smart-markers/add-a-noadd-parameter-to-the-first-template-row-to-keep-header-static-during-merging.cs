// Title: How to keep the header row static when merging cells with Aspose.Cells using ImportTableOptions and DeleteBlankOptions in C#
// AI Prompts: Import a DataTable with ShiftFirstRowDown set to false and merge a data row while keeping the header unchanged using Aspose.Cells. | Configure DeleteBlankColumns to retain only the header part of merged cells by setting MergedCellsShrinkType to KeepHeaderOnly. | Save the workbook as NoAddHeaderStatic.xlsx after preserving the static header during merge operations.
// Common Searches: Aspose.Cells keep header row static when merging cells C# | ShiftFirstRowDown false prevents header shift during ImportData Aspose.Cells | DeleteBlankColumns KeepHeaderOnly option for merged header Aspose.Cells | How to merge a data row without moving the header in Aspose.Cells .NET | Preserve header while deleting blank columns in Excel using Aspose.Cells
// Tags: ImportTableOptions ShiftFirstRowDown false | DeleteBlankOptions KeepHeaderOnly | Aspose.Cells static header during merge | merge cells without shifting header Aspose.Cells | preserve header row while deleting blank columns Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// Demonstrates importing a DataTable with the header row fixed, merging a data row, deleting blank columns while keeping only the header part of merged cells, and saving the workbook as NoAddHeaderStatic.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare a DataTable with a header row and some data rows
        DataTable table = new DataTable();
        table.Columns.Add("Header");
        table.Columns.Add("Value");
        table.Rows.Add("Header1", "Value1");
        table.Rows.Add("Header2", "Value2");
        table.Rows.Add("Header3", "Value3");

        // Import the DataTable into the worksheet.
        // ShiftFirstRowDown = false ensures the first (header) row stays at its original position
        // and is not shifted down when rows are inserted.
        ImportTableOptions importOptions = new ImportTableOptions
        {
            ShiftFirstRowDown = false,
            IsFieldNameShown = true // keep the header row visible
        };
        cells.ImportData(table, 0, 0, importOptions);

        // Merge cells in the second data row (row index 1) across two columns.
        // This demonstrates a merge operation where we want the header to remain static.
        cells.Merge(1, 0, 1, 2);

        // When performing blank column deletion, keep only the header part of merged areas.
        DeleteBlankOptions deleteOptions = new DeleteBlankOptions
        {
            MergedCellsShrinkType = MergedCellsShrinkType.KeepHeaderOnly
        };
        cells.DeleteBlankColumns(deleteOptions);

        // Save the workbook to a file.
        workbook.Save("NoAddHeaderStatic.xlsx");
    }
}
