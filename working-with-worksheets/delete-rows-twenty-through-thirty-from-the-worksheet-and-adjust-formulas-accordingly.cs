// Title: Delete rows 20‑30 from an Excel worksheet and automatically adjust formulas using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that removes rows 20 through 30 from the first worksheet and refreshes all dependent formulas. | Show how to use Aspose.Cells Cells.DeleteRows to delete a specific row range while preserving formula references in a .NET workbook. | Provide a complete example that loads an .xlsx file, deletes rows 20‑30, and saves the updated file with formulas recalculated using Aspose.Cells.
// Common Searches: Aspose.Cells C# delete rows 20 to 30 and keep formulas intact | How to remove a block of rows in an Excel file with Aspose.Cells while updating formulas | C# example for deleting rows 20‑30 in a worksheet using Aspose.Cells and preserving formula references | Delete multiple rows in .xlsx using Aspose.Cells .NET and auto‑recalculate dependent formulas
// Tags: Aspose.Cells Cells.DeleteRows row range removal | C# delete Excel rows with formula update | Aspose.Cells preserve formulas after row deletion | Excel .xlsx row block removal using Aspose.Cells .NET | automatic formula recalculation Aspose.Cells

using Aspose.Cells;

// The code loads input.xlsx, deletes rows 20‑30 (zero‑based indices 19‑29) from the first worksheet with Aspose.Cells, which automatically updates any affected formulas, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the target worksheet (first sheet in this example)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete rows 20 through 30.
        // Aspose.Cells uses zero‑based indexing, so row 20 is index 19.
        // Total rows to delete = 30 - 20 + 1 = 11.
        sheet.Cells.DeleteRows(19, 11);

        // Save the workbook (save rule)
        workbook.Save("output.xlsx");
    }
}
