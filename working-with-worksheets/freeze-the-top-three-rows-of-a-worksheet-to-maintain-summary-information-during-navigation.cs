// Title: How to freeze the top three rows of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that applies Worksheet.FreezePanes to lock the first three rows of an Excel sheet and saves the file with Aspose.Cells. | Show the implementation steps for creating a workbook, freezing rows 0‑2, and exporting the result as a frozen‑pane Excel file in C#. | Provide a concise example that demonstrates using Aspose.Cells to keep three header rows static while the user scrolls.
// Common Searches: aspnet cells freeze top three rows example c# | using Aspose.Cells to lock header rows in an Excel workbook | C# Worksheet.FreezePanes rows 0-2 Aspose.Cells tutorial | how to keep the first few rows visible when scrolling with Aspose.Cells | save Excel file with frozen rows using Aspose.Cells .NET
// Tags: freeze panes Aspose.Cells C# | lock top rows worksheet Aspose.Cells | Aspose.Cells FreezePanes method example | freeze header rows Excel .NET | save workbook with frozen panes Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet, freezes rows 0‑2 using Worksheet.FreezePanes(3,0,3,0), and saves the file as FrozenTopRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Freeze the top three rows (rows 0,1,2).
            // Parameters: split row index, split column index, number of rows to freeze, number of columns to freeze
            sheet.FreezePanes(3, 0, 3, 0);

            // Save the workbook to a file
            workbook.Save("FrozenTopRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
