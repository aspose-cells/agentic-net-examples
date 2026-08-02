using System;
using Aspose.Cells;

class JsonExportHiddenRowsColumns
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate a 5x5 grid with sample data
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                worksheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Hide a specific row and column (zero‑based indexes)
        worksheet.Cells.HideRow(1);   // Hide row 2
        worksheet.Cells.HideColumn(2); // Hide column C

        // Determine the bounding area that contains only visible rows and columns
        int startRow = int.MaxValue, endRow = int.MinValue;
        int startColumn = int.MaxValue, endColumn = int.MinValue;

        // Iterate over the used range of the worksheet
        for (int r = 0; r <= worksheet.Cells.MaxDataRow; r++)
        {
            if (worksheet.Cells.IsRowHidden(r))
                continue; // Skip hidden rows

            for (int c = 0; c <= worksheet.Cells.MaxDataColumn; c++)
            {
                if (worksheet.Cells.IsColumnHidden(c))
                    continue; // Skip hidden columns

                // Update the visible area bounds
                if (r < startRow) startRow = r;
                if (r > endRow) endRow = r;
                if (c < startColumn) startColumn = c;
                if (c > endColumn) endColumn = c;
            }
        }

        // Build the CellArea that represents only the visible portion
        CellArea visibleArea = new CellArea
        {
            StartRow = startRow,
            EndRow = endRow,
            StartColumn = startColumn,
            EndColumn = endColumn
        };

        // Configure JsonSaveOptions to use the visible area only
        JsonSaveOptions jsonOptions = new JsonSaveOptions
        {
            ExportArea = visibleArea,   // Export only visible rows/columns
            SkipEmptyRows = true,       // Optional: skip rows that become empty after hiding
            ExportEmptyCells = false,   // Optional: do not include empty cells
            HasHeaderRow = false        // Adjust according to your data layout
        };

        // Save the workbook as JSON; hidden rows/columns are excluded
        workbook.Save("output.json", jsonOptions);
    }
}