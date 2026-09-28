// Title: Freeze the first row and first column in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook with Aspose.Cells, freezes the topmost row and leftmost column, and saves it as an .xlsx file. | Show how to call Worksheet.FreezePanes to keep header rows and columns visible while scrolling in an Aspose.Cells spreadsheet. | Demonstrate setting the freeze pane parameters (rows, columns, top‑left cell) to lock the first row and column in a C# Aspose.Cells workbook.
// Common Searches: Aspose.Cells C# freeze first row and column example | how to keep Excel header row visible using Aspose.Cells FreezePanes | C# code to freeze top row and left column in an .xlsx file with Aspose.Cells | freeze panes parameters explanation Aspose.Cells .NET
// Tags: Aspose.Cells FreezePanes first row column | freeze top row left column Excel C# | worksheet freeze panes parameters Aspose.Cells | lock header rows columns Aspose.Cells .NET

using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, applies FreezePanes(1, 1, 1, 1) to lock the first row and column, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze panes at the first row and first column.
        // Parameters: totalRows to freeze, totalColumns to freeze,
        // rows and columns of the top‑left cell of the scrollable area (zero‑based).
        sheet.FreezePanes(1, 1, 1, 1);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
