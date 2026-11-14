// Title: How to freeze the first row and first column (header) in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells freeze first row and column in Excel | keep header row and column visible while scrolling with Aspose.Cells | Aspose.Cells FreezePanes example for freezing top row and left column | freeze panes at A1 using Aspose.Cells .NET API | set frozen panes programmatically in an Excel workbook with C#
// Tags: Aspose.Cells FreezePanes API C# | freeze top row and left column Excel .NET | programmatic frozen panes Excel workbook | keep header visible scrolling Aspose.Cells | Excel worksheet freeze panes C# example

using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, calls FreezePanes(1,1,1,1) to lock the first row and column, and saves the file as FrozenPane.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze the first row and first column (A1 cell)
        // Parameters: row, column, totalRows, totalColumns (zero‑based indexes)
        // Setting row = 1 and column = 1 freezes everything above row 1 and left of column 1.
        sheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook to a file
        workbook.Save("FrozenPane.xlsx");
    }
}
