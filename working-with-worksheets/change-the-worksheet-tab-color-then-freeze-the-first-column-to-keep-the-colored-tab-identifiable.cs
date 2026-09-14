// Title: Set worksheet tab color to LightBlue and freeze column A using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to change the first worksheet's tab color to LightBlue and freeze column A while leaving rows unfrozen. | Generate a .NET snippet that applies a LightBlue tab color to a worksheet and calls FreezePanes to lock the first column in an Excel file with Aspose.Cells.
// Common Searches: Aspose.Cells C# change worksheet tab color to LightBlue | Freeze first column in Excel workbook using Aspose.Cells .NET | How to apply tab color and freeze panes together with Aspose.Cells | C# Aspose.Cells example set tab color and freeze column A
// Tags: Aspose.Cells worksheet tab color C# | Aspose.Cells freeze first column C# | LightBlue tab styling Aspose.Cells | FreezePanes column A Aspose.Cells | Excel workbook layout styling Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new Workbook, accesses the first Worksheet, sets its TabColor to LightBlue, freezes column A by calling FreezePanes(0, 1, 0, 1), and saves the workbook as Output.xlsx, with exception handling.
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

            // Change the worksheet tab color (e.g., LightBlue)
            sheet.TabColor = Color.LightBlue;

            // Freeze the first column (column A)
            // row = 0 (no frozen rows), column = 1 (freeze up to column A),
            // freezedRows = 0, freezedColumns = 1 (freeze one column)
            sheet.FreezePanes(0, 1, 0, 1);

            // Save the workbook
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
