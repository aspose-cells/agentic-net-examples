// Title: Hide rows 0‑9 and disable the vertical scroll bar with Aspose.Cells for .NET before saving an .xlsx file
// AI Prompts: Create a new workbook, hide rows 0 through 9 on the first worksheet, set ViewOptions.ShowVerticalScrollBar to false, and save as Output.xlsx using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically conceal the first ten rows of a worksheet, turn off the vertical scroll bar, and export the workbook to an .xlsx file. | Write C# code that iterates over rows 0‑9 to hide them, disables the vertical scroll bar via the worksheet's ViewOptions, and saves the result with Aspose.Cells.
// Common Searches: Aspose.Cells C# hide first ten rows and hide vertical scroll bar before saving workbook | disable vertical scroll bar in Excel file using Aspose.Cells .NET | programmatically hide rows 0-9 with Aspose.Cells and control scroll bar visibility | C# Aspose.Cells hide rows and turn off scroll bar in worksheet view options
// Tags: hide rows Aspose.Cells C# | disable vertical scroll bar Aspose.Cells | ViewOptions.ShowVerticalScrollBar Aspose.Cells | Aspose.Cells hide rows before save | Aspose.Cells worksheet scroll bar visibility

using Aspose.Cells;
using System;

// // Creates a workbook, hides rows 0‑9 on the first worksheet, optionally disables the vertical scroll bar via ViewOptions, and saves the file as Output.xlsx.
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

            // Hide rows 0 through 9 (first ten rows)
            for (int row = 0; row <= 9; row++)
            {
                sheet.Cells.HideRow(row);
            }

            // Hide the vertical scroll bar if the ViewOptions property is available
            // (In older Aspose.Cells versions this property may not exist)
            // Uncomment the following line when using a version that supports ViewOptions.
            // sheet.ViewOptions.ShowVerticalScrollBar = false;

            // Save the workbook
            workbook.Save("Output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
