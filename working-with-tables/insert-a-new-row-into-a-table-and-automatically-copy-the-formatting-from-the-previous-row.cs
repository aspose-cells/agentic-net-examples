// Title: Insert a new row into an Excel worksheet and copy the previous row’s formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# insert row and keep original row formatting | how to copy style from previous row after inserting rows with Aspose.Cells | preserve Excel table row formatting when adding a new row in .NET | C# code to insert row at specific index and duplicate row style using Aspose.Cells | Aspose.Cells insert rows with formatting inheritance example
// Tags: Aspose.Cells insert rows with style inheritance | C# copy row formatting after row insertion | Excel worksheet row formatting preservation using Aspose | automate row style duplication Aspose.Cells | preserve table row style .NET Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing workbook, inserts a new row at a given index while shifting existing rows down, then copies the style from the row above to each cell of the inserted row, and finally saves the workbook with the retained formatting.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Row index where the new row will be inserted (0‑based). 
        // For example, insert a row at Excel row 6 (index 5).
        int insertRowIndex = 5;

        // Insert one new row at the specified index.
        // The third parameter (true) shifts existing rows down.
        cells.InsertRows(insertRowIndex, 1, true);

        // Copy formatting from the row above the inserted row.
        int sourceRowIndex = insertRowIndex - 1;
        int maxColumn = cells.MaxColumn; // highest column used in the sheet

        for (int col = 0; col <= maxColumn; col++)
        {
            // Get the style of the source cell.
            Style sourceStyle = cells[sourceRowIndex, col].GetStyle();

            // Apply the same style to the newly inserted cell.
            cells[insertRowIndex, col].SetStyle(sourceStyle);
        }

        // Save the modified workbook.
        workbook.Save("output.xlsx");
    }
}
