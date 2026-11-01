// Title: How to set column width to 20 characters for columns D‑F using Range.EntireColumn in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set width of columns D, E, and F to 20 characters | Using Range.EntireColumn to change column widths in a .NET workbook | Programmatically adjust multiple column widths with Aspose.Cells API | Set column width for a range of columns in Aspose.Cells C# example | How to apply the same column width to columns D through F in Aspose.Cells
// Tags: Range.EntireColumn column width adjustment | Aspose.Cells set multiple column widths | C# set column width for specific columns | column width adjustment Aspose.Cells | define column range D-F Aspose.Cells

// Create a new workbook
var workbook = new Aspose.Cells.Workbook();

// Access the first worksheet
var worksheet = workbook.Worksheets[0];

// Define a range that covers columns D through F (row 1 is arbitrary, only columns matter)
Aspose.Cells.Range range = worksheet.Cells.CreateRange("D1:F1");

// Use EntireColumn to get the full columns of the defined range
Aspose.Cells.Range entireColumns = range.EntireColumn;

// Set each column width to 20 characters
for (int i = 0; i < entireColumns.ColumnCount; i++)
{
    // Calculate the actual column index (0‑based)
    int columnIndex = entireColumns.FirstColumn + i;

    // Apply the width
    worksheet.Cells.SetColumnWidth(columnIndex, 20);
}

// Save the workbook
workbook.Save("Result.xlsx");
