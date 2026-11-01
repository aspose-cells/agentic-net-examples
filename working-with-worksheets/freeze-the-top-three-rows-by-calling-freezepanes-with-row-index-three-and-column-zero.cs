// Title: How to freeze the first three rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, freezes rows 1‑3 on the first worksheet, and saves it as an .xlsx file. | Show how to call Worksheet.FreezePanes to lock the top three rows (row index 3, column index 0) in a newly created Excel file using Aspose.Cells.
// Common Searches: Aspose.Cells C# freeze top three rows of worksheet | C# FreezePanes method to lock first three rows in Excel file | How to use Aspose.Cells to freeze rows when creating a workbook programmatically | Freeze panes rows index 3 column 0 Aspose.Cells .NET example
// Tags: Worksheet.FreezePanes usage Aspose.Cells | freeze rows programmatically Aspose.Cells | Excel workbook with frozen header Aspose | C# Aspose.Cells save frozen pane file | lock first rows Excel via Aspose API

using Aspose.Cells;
using System;

// Creates a new workbook, accesses the first worksheet, freezes the first three rows using Worksheet.FreezePanes(3,0,3,0), and saves the file as FrozenRows.xlsx.
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

            // Freeze the top three rows (row index 3, column index 0)
            // The overload requires row, column, totalRows, totalColumns
            sheet.FreezePanes(3, 0, 3, 0);

            // Save the workbook to a file
            workbook.Save("FrozenRows.xlsx");
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
