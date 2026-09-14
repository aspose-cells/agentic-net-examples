// Title: C# guide to freezing the top row and leftmost column simultaneously with Aspose.Cells FreezePanes
// AI Prompts: Generate C# code that creates a workbook, accesses the first worksheet, and applies FreezePanes to lock the top row and left column at the same time. | Explain the required parameter values for the FreezePanes method to freeze row 1 and column 1 in Aspose.Cells. | Provide a complete C# snippet that saves the workbook after freezing the first row and column using Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze top row and left column together | How to use FreezePanes to lock row 1 and column 1 in a .NET Excel workbook | C# example for freezing pane at A1 with Aspose.Cells library | Simultaneous freeze of first row and first column using Aspose.Cells in C# | Freeze first row and column in Excel file with Aspose.Cells API
// Tags: Aspose.Cells FreezePanes first row column C# | freeze top row left column Aspose.Cells | C# Excel FreezePanes example | Aspose.Cells workbook freeze pane A1 | lock first row and column worksheet Aspose

using System;
using Aspose.Cells;

// Creates a new Workbook, retrieves the first Worksheet, calls sheet.FreezePanes(1,1,1,1) to lock the top row and leftmost column, and saves the file as FrozenPane.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the first row and first column simultaneously.
            // Parameters: row index (1) and column index (1) are zero‑based offsets for the pane below/right of the frozen area.
            // totalRows = 1, totalColumns = 1 specify that one row and one column are frozen.
            sheet.FreezePanes(1, 1, 1, 1);

            // Save the workbook
            workbook.Save("FrozenPane.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
