// Title: Unprotect an Excel worksheet while retaining formatting protection and set a cell formula using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file with Aspose.Cells, remove worksheet protection while keeping its formatting protection, assign the formula '=SUM(B1:B10)' to cell A1, and save the workbook as a new file. | Write C# code that opens a workbook, calls Unprotect on the first worksheet (preserving formatting settings), updates a specific cell's formula, and writes the updated file using Aspose.Cells.
// Common Searches: Aspose.Cells C# remove worksheet protection keep formatting settings | how to change a cell formula after unprotecting a sheet with Aspose.Cells | preserve formatting permissions when disabling sheet protection in .NET | unprotect Excel worksheet without password using Aspose.Cells example | update A1 to SUM(B1:B10) after sheet unprotection Aspose.Cells
// Tags: worksheet unprotect retain formatting Aspose.Cells | set cell formula after sheet unprotection C# | load workbook modify formula Aspose.Cells | preserve sheet formatting permissions .NET | Excel sheet unprotect no password Aspose.Cells

using Aspose.Cells;

// Loads input.xlsx, unprotects the first worksheet while preserving its formatting protection, sets A1 formula to =SUM(B1:B10), and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or specify by name)
        Worksheet sheet = workbook.Worksheets[0];

        // Preserve current formatting permissions (they remain in the Protection object)
        // Unprotect the worksheet (provide password if it was protected with one)
        sheet.Unprotect(); // No password argument assumes no password; use sheet.Unprotect("password") if needed

        // Update a cell formula (example: set A1 to sum of B1:B10)
        Cell targetCell = sheet.Cells["A1"];
        targetCell.Formula = "=SUM(B1:B10)";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
