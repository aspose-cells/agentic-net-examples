// Title: Read the formula of cell E3 in the second worksheet of an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells in C# to obtain the Formula property of cell E3 on the workbook's second worksheet and output it. | Load an .xlsx file, locate the worksheet at index 1, and retrieve the formula string from cell E3 using Aspose.Cells. | Extract and display the Excel formula stored in cell E3 before deleting the corresponding worksheet with Aspose.Cells API.
// Common Searches: Aspose.Cells C# retrieve formula from cell E3 on second sheet | how to get Excel cell formula before removing worksheet using .NET | read formula of a specific cell in a particular worksheet with Aspose.Cells | C# code to access formula property of a cell in workbook's second worksheet | extract formula string from Excel file without opening UI Aspose.Cells
// Tags: Aspose.Cells read cell formula .NET | C# access second worksheet cell E3 | Formula property usage Aspose.Cells | extract formula from cell before sheet removal | workbook worksheet index cell access Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook, selects the second worksheet (index 1), accesses cell E3, reads its Formula property (empty if no formula is present), and prints the formula string to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the second worksheet (index 1)
        Worksheet secondSheet = workbook.Worksheets[1];

        // Get the cell E3 from that worksheet
        Cell targetCell = secondSheet.Cells["E3"];

        // Retrieve the formula string (empty if the cell has no formula)
        string formula = targetCell.Formula;

        // Output the formula to the console
        Console.WriteLine($"Formula in worksheet '{secondSheet.Name}' cell E3: {formula}");
    }
}
