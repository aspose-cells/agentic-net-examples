// Title: Insert a new row in an Excel worksheet and copy the previous row’s formatting with Aspose.Cells for .NET (C#)
// AI Prompts: Add a row after a specific index and duplicate the style of the preceding row using Aspose.Cells in C#. | Copy only the formatting (fonts, borders, fills) from an existing row to a newly inserted row while leaving cell values empty with Aspose.Cells.
// Common Searches: Aspose.Cells C# insert row after row 5 and keep original formatting | How to copy row style without values when adding a new row in Excel using Aspose.Cells | C# code to insert a row and apply previous row’s formatting in an .xlsx file with Aspose.Cells
// Tags: insert row with Aspose.Cells C# | copy row formatting Aspose.Cells | preserve Excel row style on insertion | Aspose.Cells worksheet row formatting

using Aspose.Cells;

// Loads input.xlsx, inserts a new row after row 5, copies the style from the preceding row across all used columns, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Row after which the new row will be inserted (zero‑based index)
        int insertAfterRow = 5;
        int newRowIndex = insertAfterRow + 1;

        // Insert a single row at the desired position and shift existing rows down
        sheet.Cells.InsertRows(newRowIndex, 1, true);

        // Determine the last used column to know how many cells to copy
        int maxColumn = sheet.Cells.MaxColumn;

        // Copy formatting from the previous row to the newly inserted row
        for (int col = 0; col <= maxColumn; col++)
        {
            Cell sourceCell = sheet.Cells[insertAfterRow, col];
            Cell targetCell = sheet.Cells[newRowIndex, col];

            // Transfer the style (font, borders, fill, etc.)
            targetCell.SetStyle(sourceCell.GetStyle());

            // If you want the new row to be empty, do NOT copy the value.
            // To copy the value as well, uncomment the following line:
            // targetCell.PutValue(sourceCell.Value);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
