// Title: Read the updated formula in cell E3 of the second worksheet after deleting two rows with reference updates using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to delete rows 1 and 2 (updating references) and then retrieve the formula from cell E3 on the second worksheet. | Show C# code that removes specific rows while preserving formulas and reads the resulting formula in E3 of the second sheet.
// Common Searches: Aspose.Cells C# get formula from E3 after deleting rows with updateReference true | How to preserve formulas when deleting rows in a specific worksheet using Aspose.Cells | Read cell formula after row deletions in second worksheet Aspose.Cells .NET | Delete first two rows and retrieve updated formula in Excel file using Aspose.Cells
// Tags: row deletion preserving formulas Aspose.Cells | read cell formula after row removal .NET | second worksheet formula extraction Aspose.Cells | Aspose.Cells DeleteRow updateReference usage | C# Excel workbook manipulate rows and formulas

using System;
using Aspose.Cells;

// Loads an Excel workbook, accesses the second worksheet, deletes the first row and then the next row while updating cell references, reads the formula from cell E3, prints it, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Verify that there is a second worksheet
        if (workbook.Worksheets.Count < 2)
        {
            Console.WriteLine("The workbook must contain at least two worksheets.");
            return;
        }

        // Access the second worksheet (index 1)
        Worksheet sheet = workbook.Worksheets[1];
        Cells cells = sheet.Cells;

        // First deletion: delete the first row (index 0) and update references
        cells.DeleteRow(0, true);

        // Second deletion: delete what is now the second row (originally row 2)
        cells.DeleteRow(1, true);

        // After the deletions, read the formula from cell E3 (row 2, column 4)
        string formula = cells["E3"].Formula;

        Console.WriteLine($"Formula in E3 after the second deletion: {formula}");

        // Save the workbook (optional, depending on whether you need to persist changes)
        workbook.Save("output.xlsx");
    }
}
