// Title: Freeze the First Row and Column in Excel with Aspose.Cells for .NET (C#)
// Description: Learn how to use Aspose.Cells for .NET to freeze the topmost row and leftmost column of a worksheet. The example creates a workbook, calls FreezePanes(1,1,1,1) to keep headers visible while scrolling, and saves the file as FreezePanesFirstRowColumn.xlsx.
// Keywords: Aspose.Cells FreezePanes C# | freeze top row Excel C# | freeze first column Aspose.Cells | keep headers visible Excel | Aspose.Cells worksheet freeze panes | C# Excel freeze panes API | Aspose.Cells FreezePanes method
// Common Searches: Aspose.Cells how to freeze first row and column | C# FreezePanes example for Excel | keep header row static while scrolling Aspose.Cells | freeze panes in Excel using Aspose.Cells for .NET | Aspose.Cells FreezePanes parameters explained
// Developer Intent: Apply FreezePanes to lock the first row and first column so header labels stay in view during scrolling.
// Use Cases: Create reporting workbooks where column and row headers remain visible in large data tables. | Export dynamic datasets to Excel with frozen panes to improve readability for end‑users. | Prepare a template workbook with frozen header rows/columns before populating it with data.
// AI Prompts: Generate C# code to freeze multiple rows and columns with Aspose.Cells. | Show how to unfreeze panes and reset scrolling in an existing Excel file using Aspose.Cells for .NET. | Explain each parameter of the FreezePanes method and how they define the frozen area.

using System;
using Aspose.Cells;

// Learn how to use Aspose.Cells for .NET to freeze the topmost row and leftmost column of a worksheet. The example creates a workbook, calls FreezePanes(1,1,1,1) to keep headers visible while scrolling, and saves the file as FreezePanesFirstRowColumn.xlsx.
class FreezePanesExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze the first row and first column (row index 1, column index 1)
        sheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook
        workbook.Save("FreezePanesFirstRowColumn.xlsx");
    }
}
