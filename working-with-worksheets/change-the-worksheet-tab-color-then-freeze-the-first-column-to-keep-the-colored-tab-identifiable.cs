// Title: Set worksheet tab color to LightBlue and freeze column A using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that changes a worksheet's tab color to LightBlue and freezes the first column (column A) with Aspose.Cells. | Show how to apply a tab color and create a freeze pane on column A in an Aspose.Cells workbook using .NET.
// Common Searches: Aspose.Cells C# how to change worksheet tab color and freeze column A | C# code to set tab color LightBlue and freeze first column in Excel file with Aspose.Cells | freeze panes on first column while preserving tab color using Aspose.Cells .NET | set worksheet tab color and freeze column A programmatically Aspose.Cells
// Tags: set worksheet tab color Aspose.Cells | freeze first column Aspose.Cells | Aspose.Cells FreezePanes method | Aspose.Cells tab color customization | C# Aspose.Cells workbook formatting

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, sets its TabColor to LightBlue, freezes column A by calling FreezePanes(0, 1, 0, 1), and saves the workbook as Result.xlsx, with exception handling.
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

            // Change the worksheet tab color (e.g., LightBlue)
            sheet.TabColor = Color.LightBlue;

            // Freeze the first column (no rows, one column)
            // row = 0 (first row), column = 1 (second column = B), totalRows = 0, totalColumns = 1
            sheet.FreezePanes(0, 1, 0, 1);

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
