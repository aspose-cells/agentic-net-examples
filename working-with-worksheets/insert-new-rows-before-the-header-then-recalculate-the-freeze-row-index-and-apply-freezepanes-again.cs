// Title: Insert rows before a header row and adjust FreezePanes in Aspose.Cells for .NET
// Common Searches: Aspose.Cells insert rows before a specific row and keep freeze panes | C# how to shift freeze pane row after inserting rows in Excel using Aspose.Cells | recalculate FreezePanes row index after inserting rows with Aspose.Cells .NET | update worksheet freeze panes when adding rows before header in Aspose.Cells | adjust freeze pane position after row insertion in C# Aspose.Cells workbook
// Tags: Aspose.Cells insert rows before header | Aspose.Cells recalculate FreezePanes after insertion | C# update freeze pane row index | Excel worksheet row insertion with Aspose.Cells | adjust worksheet freeze panes programmatically

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Get the first worksheet (adjust index if needed)
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Define the header row index (0‑based). Rows will be inserted before this row.
int headerRowIndex = 5;          // example: header is at row 6 in Excel (index 5)
int rowsToInsert = 2;            // number of rows to insert

// Insert rows before the header
sheet.Cells.InsertRows(headerRowIndex, rowsToInsert);

// -------------------------------------------------------------------
// Recalculate FreezePanes
// -------------------------------------------------------------------
// Assume the worksheet currently has freeze panes set at a specific row.
// Retrieve the existing freeze row (if any). Aspose.Cells does not expose
// a direct getter, so we store the original split values before insertion.
// For demonstration, we assume the original freeze row was 4 (row 5 in Excel).
int originalFreezeRow = 4;       // replace with actual value if known

// If the original freeze row is at or below the insertion point, shift it down.
int newFreezeRow = originalFreezeRow;
if (originalFreezeRow >= headerRowIndex)
{
    newFreezeRow += rowsToInsert;
}

// Apply FreezePanes again with the updated row index.
// Freeze the rows above 'newFreezeRow' and keep columns unchanged (column split = 0).
sheet.FreezePanes(newFreezeRow, 0, newFreezeRow, 0);

// -------------------------------------------------------------------
// Save the modified workbook
// -------------------------------------------------------------------
workbook.Save("output.xlsx");
