// Title: How to unhide row 10 in the active worksheet of an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, reveals a hidden row in the active sheet, and saves the workbook. | Show the steps to change the IsHidden property of a specific row in Aspose.Cells and persist the changes to a new file. | Provide a minimal Aspose.Cells example that loads a workbook, accesses the active worksheet, makes a hidden row visible, and writes the result.
// Common Searches: Aspose.Cells C# unhide hidden row in active worksheet | set Row.IsHidden false for a specific row using Aspose.Cells .NET | load Excel file, modify row visibility, and save with Aspose.Cells | how to make row 10 visible in an .xlsx file using Aspose.Cells | C# example for changing row visibility in Aspose.Cells workbook
// Tags: unhide row Aspose.Cells .NET | Row.IsHidden property Aspose.Cells | active worksheet row manipulation Aspose.Cells | load and save Excel workbook Aspose.Cells | modify row visibility C# Aspose.Cells

using Aspose.Cells;
using System;

// The program loads 'input.xlsx' with Aspose.Cells, accesses the active worksheet, unhides row 10 by setting its IsHidden property to false, and saves the updated workbook as 'output.xlsx'.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the active worksheet
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Row indices are zero‑based; row 10 is index 9
        Row row10 = activeSheet.Cells.Rows[9];

        // Unhide the row
        row10.IsHidden = false;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
