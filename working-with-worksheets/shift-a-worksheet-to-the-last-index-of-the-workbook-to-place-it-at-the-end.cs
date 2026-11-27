// Title: Move a worksheet to the last position in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# move worksheet to the end of an Excel workbook | How to reorder worksheets so a specific sheet becomes last using Aspose.Cells | C# example for moving a worksheet to the final index with Aspose.Cells MoveTo | Place first worksheet at the last index in a .xlsx file using Aspose.Cells .NET
// Tags: move worksheet to last index Aspose.Cells | worksheet reordering C# Aspose.Cells | Aspose.Cells MoveTo method usage | place worksheet at end of Excel workbook .NET | shift worksheet position Aspose.Cells example

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Identify the worksheet you want to move (by name or index)
// Example: move the first worksheet
Aspose.Cells.Worksheet sheetToMove = workbook.Worksheets[0];

// Determine the last possible index in the workbook
int lastIndex = workbook.Worksheets.Count - 1;

// If the worksheet is not already at the last position, move it there
if (sheetToMove.Index != lastIndex)
{
    sheetToMove.MoveTo(lastIndex);
}

// Save the modified workbook
workbook.Save("output.xlsx");
