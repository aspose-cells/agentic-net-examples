// Title: Read and clear the formula of cell E3 in the second worksheet with Aspose.Cells for .NET
// Description: Loads an Excel file, accesses the second worksheet (index 1), retrieves the Formula property of cell E3, prints it, empties the formula, and saves the workbook. Demonstrates how to get a cell's formula before removal using Aspose.Cells in C#.
// Keywords: Aspose.Cells | C# | read cell formula | clear cell formula | second worksheet | Excel formula retrieval | Cell.Formula | Workbook manipulation | retrieve formula before deletion
// Common Searches: Aspose.Cells get formula of cell E3 in second sheet | C# read and delete Excel cell formula with Aspose | How to retrieve a formula before clearing it using Aspose.Cells | Access second worksheet cell formula Aspose.Cells .NET | Extract and remove formula from specific Excel cell
// Developer Intent: Extract the formula from E3 on the second worksheet, display it, then clear the formula and save the file.
// Use Cases: Log existing formulas before bulk‑clearing a sheet | Validate formulas prior to data transformation or cleanup | Export formulas for documentation or migration to another system
// AI Prompts: Write C# code with Aspose.Cells that reads the formula of cell E3 on the second worksheet, outputs it, clears the formula, and saves the workbook. | Show how to capture a cell's formula, write it to a log file, then delete the formula using Aspose.Cells. | Explain error handling when cell E3 has no formula while using Aspose.Cells to read and clear it.

using System;
using Aspose.Cells;

// Loads an Excel file, accesses the second worksheet (index 1), retrieves the Formula property of cell E3, prints it, empties the formula, and saves the workbook. Demonstrates how to get a cell's formula before removal using Aspose.Cells in C#.
class RetrieveFormulaBeforeDeletion
{
    static void Main()
    {
        // Load the workbook from a file
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the second worksheet (index 1)
        Worksheet secondSheet = workbook.Worksheets[1];

        // Retrieve the formula from cell E3
        Cell cellE3 = secondSheet.Cells["E3"];
        string formula = cellE3.Formula;

        // Display the retrieved formula
        Console.WriteLine($"Formula in worksheet '{secondSheet.Name}' cell E3 before deletion: {formula}");

        // Delete the cell content (clears both value and formula)
        cellE3.Formula = string.Empty;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
