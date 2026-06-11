using System;
using Aspose.Cells;

class FreezePanesPersistenceTest
{
    static void Main()
    {
        // ---------- Create a new workbook and freeze the first column ----------
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Freeze panes at cell B1 (row index 0, column index 1) with 0 frozen rows and 1 frozen column.
        // This effectively freezes column A.
        worksheet.FreezePanes(0, 1, 0, 1);

        // ---------- Save the workbook ----------
        string filePath = "FreezePanesTest.xlsx";
        workbook.Save(filePath);

        // ---------- Reload the workbook ----------
        Workbook reloadedWorkbook = new Workbook(filePath);
        Worksheet reloadedWorksheet = reloadedWorkbook.Worksheets[0];

        // ---------- Verify that the freeze panes setting persisted ----------
        bool hasFreeze = reloadedWorksheet.GetFreezedPanes(
            out int frozenRow,
            out int frozenColumn,
            out int frozenRows,
            out int frozenColumns);

        if (!hasFreeze)
        {
            throw new Exception("Freeze panes were not persisted after reload.");
        }

        // Expect 0 frozen rows and 1 frozen column (first column frozen)
        if (frozenRows != 0 || frozenColumns != 1)
        {
            throw new Exception($"Unexpected freeze settings after reload. FrozenRows={frozenRows}, FrozenColumns={frozenColumns}");
        }

        Console.WriteLine("Freeze panes persisted correctly after saving and reloading.");
    }
}