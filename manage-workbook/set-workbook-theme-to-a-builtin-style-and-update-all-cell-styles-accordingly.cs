using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add some sample data
        sheet.Cells["A1"].PutValue("Header");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B1"].PutValue("Another");
        sheet.Cells["B2"].PutValue(456);

        // -----------------------------------------------------------------
        // Set the workbook theme by applying a built‑in style (e.g., Good)
        // -----------------------------------------------------------------
        // Create the built‑in style object
        Style builtinStyle = workbook.CreateBuiltinStyle(BuiltinStyleType.Good);

        // Make this style the default style for the whole workbook
        workbook.DefaultStyle = builtinStyle;

        // ---------------------------------------------------------------
        // Update all existing cells to use the new default style
        // ---------------------------------------------------------------
        // Get the used range of the worksheet
        var usedRange = sheet.Cells.MaxDisplayRange;

        for (int row = usedRange.FirstRow; row < usedRange.FirstRow + usedRange.RowCount; row++)
        {
            for (int col = usedRange.FirstColumn; col < usedRange.FirstColumn + usedRange.ColumnCount; col++)
            {
                Cell cell = sheet.Cells[row, col];
                // Apply the default style while preserving the cell value
                cell.SetStyle(workbook.DefaultStyle);
            }
        }

        // Save the workbook
        workbook.Save("BuiltInThemeDemo.xlsx");
    }
}