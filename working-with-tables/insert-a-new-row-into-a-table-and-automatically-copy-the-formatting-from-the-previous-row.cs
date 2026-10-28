// Title: Insert a row into an Excel table and copy the previous row’s formatting with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# insert a row after a specific row while preserving the table's formatting | Copy formatting from the previous row to a newly added row with Aspose.Cells | Add a row to an Excel table programmatically and retain its style using Aspose.Cells .NET | How to duplicate the style of a preceding row when inserting rows via Aspose.Cells | Maintain Excel table row formatting when inserting rows through Aspose.Cells API
// Tags: Aspose.Cells row insertion preserving style | C# replicate Excel row formatting using Aspose.Cells | keep table formatting after adding rows Aspose.Cells | insert row at index Aspose.Cells .NET API | apply previous row style to new row Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing workbook, defines the table range, inserts a new row after a specified index, copies the style from each cell of the preceding row across the table columns to the newly inserted cells, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Define the table boundaries (example: A1:D10)
        int tableStartRow = 0;          // Row 1 (zero‑based)
        int tableEndRow = 9;            // Row 10
        int tableStartColumn = 0;       // Column A
        int tableEndColumn = 3;         // Column D

        // Row after which the new row will be inserted (e.g., after row 5)
        int insertAfterRow = 4;         // Row 5 (zero‑based)
        int newRowIndex = insertAfterRow + 1;

        // Insert a new row at the calculated position
        sheet.Cells.InsertRows(newRowIndex, 1);

        // Copy formatting from the previous row (insertAfterRow) to the new row
        for (int col = tableStartColumn; col <= tableEndColumn; col++)
        {
            // Retrieve the style of the source cell
            Style sourceStyle = sheet.Cells[insertAfterRow, col].GetStyle();

            // Apply the retrieved style to the newly inserted cell
            sheet.Cells[newRowIndex, col].SetStyle(sourceStyle);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
