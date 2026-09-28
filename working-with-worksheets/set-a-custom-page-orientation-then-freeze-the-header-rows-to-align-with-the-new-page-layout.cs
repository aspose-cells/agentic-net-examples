// Title: Set worksheet to landscape orientation and freeze the top header row using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, sets the first worksheet to landscape page orientation, freezes the first row, and saves the file as .xlsx. | Write a C# snippet that applies a custom page layout (landscape) to a worksheet, then locks the header row with FreezePanes before exporting the workbook.
// Common Searches: Aspose.Cells C# set worksheet page orientation to landscape and freeze header row | how to freeze top row after changing page setup with Aspose.Cells .NET | C# example for landscape orientation and frozen panes using Aspose.Cells
// Tags: Aspose.Cells set worksheet landscape orientation | Aspose.Cells freeze first row panes | Aspose.Cells page setup orientation C# | Aspose.Cells freeze panes header row .NET | Aspose.Cells export workbook to xlsx custom layout

using Aspose.Cells;
using System;

// Creates a new workbook, changes the first worksheet to landscape orientation, freezes the first row, and saves the result as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom page orientation (Landscape)
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Freeze the header row (first row)
            // FreezePanes(row, column, totalRows, totalColumns) freezes rows above and columns left of the specified cell.
            sheet.FreezePanes(1, 0, 1, 0); // Freeze first row

            // Save the workbook
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
