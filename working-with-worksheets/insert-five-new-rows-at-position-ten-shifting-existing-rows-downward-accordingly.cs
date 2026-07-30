// Title: Insert 5 Rows at Row 10 in an Aspose.Cells Worksheet (C#) – Shift Existing Rows Down
// Description: C# example that creates a workbook, fills column A with 15 values, then calls Cells.InsertRows(9, 5) to add five rows at the 10th position (zero‑based index). The original rows are automatically shifted down, optional data is written to column B, and the file is saved as InsertRowsAtTen.xlsx. Ideal for Excel automation with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | InsertRows | add rows worksheet | shift rows down | Excel automation | zero based index | insert multiple rows | sample code | GitHub example
// Common Searches: Aspose.Cells insert rows at specific index | C# insert rows row 10 Aspose.Cells | how to shift rows down after inserting in Aspose.Cells | InsertRows example .NET | add five rows to Excel worksheet programmatically
// Developer Intent: Add five new rows at row 10 and move existing rows downward in an Aspose.Cells worksheet.
// Use Cases: Insert a header block before existing data in a generated report. | Create space for new records in a table without overwriting current rows. | Expand a worksheet dynamically while preserving formulas and formatting.
// AI Prompts: Generate C# code that inserts N rows at a given zero‑based index in an Aspose.Cells worksheet and keeps existing formulas intact. | Show how to copy the formatting of the preceding row to newly inserted rows after calling InsertRows in Aspose.Cells. | Write a loop that fills newly inserted rows with values from a List<string> using Aspose.Cells.

using System;
using Aspose.Cells;

// C# example that creates a workbook, fills column A with 15 values, then calls Cells.InsertRows(9, 5) to add five rows at the 10th position (zero‑based index). The original rows are automatically shifted down, optional data is written to column B, and the file is saved as InsertRowsAtTen.xlsx. Ideal for Excel automation with Aspose.Cells for .NET.
class InsertRowsExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill some sample data to demonstrate the shift after insertion
        for (int i = 0; i < 15; i++)
        {
            cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Insert five rows at position ten.
        // Row index is zero‑based, so position ten corresponds to index 9.
        cells.InsertRows(9, 5);

        // Optionally add data to the newly inserted rows (columns B)
        for (int i = 9; i < 14; i++)
        {
            cells[i, 1].PutValue("Inserted");
        }

        // Save the workbook
        workbook.Save("InsertRowsAtTen.xlsx");
    }
}
