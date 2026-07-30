// Title: Reusable C# FreezePanes Helper for Aspose.Cells Worksheets
// Description: Shows how to encapsulate Aspose.Cells' FreezePanes call in a single C# method that takes row and column indexes, freezes the rows above and columns left of the specified cell, and then saves the workbook.
// Keywords: Aspose.Cells | C# | FreezePanes | worksheet helper | Excel freeze pane | row index | column index | utility method | .NET | Excel automation
// Common Searches: Aspose.Cells freeze panes reusable method | C# helper for FreezePanes in Excel | how to freeze rows and columns with Aspose.Cells | wrap FreezePanes in a function .NET | apply same freeze pane to multiple worksheets
// Developer Intent: Create a single method that abstracts the FreezePanes operation, allowing the caller to specify the target row and column.
// Use Cases: Freeze the header rows and side columns of a generated report with one call. | Apply identical freeze settings to several sheets by passing different coordinates. | Provide a UI‑driven feature where users pick a cell and the workbook freezes panes accordingly before export.
// AI Prompts: Write a static C# class named ExcelPaneHelper with a method FreezePane(Worksheet ws, int row, int col) for Aspose.Cells. | Show code that iterates over three worksheets and calls the helper with distinct row/column values. | Explain why FreezePanes receives the row and column parameters twice and how that determines the frozen area.

using Aspose.Cells;

// Shows how to encapsulate Aspose.Cells' FreezePanes call in a single C# method that takes row and column indexes, freezes the rows above and columns left of the specified cell, and then saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze panes at row 3, column 3 (zero‑based indexes)
        FreezePane(sheet, 3, 3);

        // Save the workbook (save rule)
        workbook.Save("FreezePanesReusableDemo.xlsx");
    }

    // Reusable method that encapsulates FreezePanes logic
    static void FreezePane(Worksheet worksheet, int row, int column)
    {
        // Freeze the pane at the specified cell.
        // The last two parameters define how many rows and columns are frozen.
        // Using the same values as the row and column indexes freezes all rows above
        // and all columns to the left of the specified cell.
        worksheet.FreezePanes(row, column, row, column);
    }
}
