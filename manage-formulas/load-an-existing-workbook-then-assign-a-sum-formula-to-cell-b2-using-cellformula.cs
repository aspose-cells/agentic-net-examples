// Title: Assign a SUM formula to cell B2 in an existing .xlsx workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing Excel file with Aspose.Cells, writes the formula "=SUM(A1:A10)" into cell B2 of the first worksheet, and saves the result as a new file. | Create a C# snippet using Aspose.Cells that loads a workbook, determines the last used row in column A, sets a dynamic SUM formula covering A1 to that row in cell B2, and writes the workbook back. | Write C# code with Aspose.Cells that includes try‑catch error handling while loading a workbook, assigning a formula to a specific cell, and saving the updated workbook.
// Common Searches: C# Aspose.Cells how to set a formula in a specific cell of an existing workbook | using Aspose.Cells to add a SUM(A1:A10) formula to cell B2 in .xlsx | load Excel file with Aspose.Cells and write formula programmatically in .NET | Aspose.Cells set formula and save workbook without losing existing data | example code for assigning formulas to cells with Aspose.Cells for .NET
// Tags: Aspose.Cells assign SUM formula C# | load existing .xlsx workbook Aspose.Cells | assign formula to B2 using Aspose.Cells | save workbook after formula update Aspose | dynamic range SUM formula Aspose.Cells C#

using Aspose.Cells;
using System;

// Loads input.xlsx, writes the formula "=SUM(A1:A10)" into cell B2 of the first worksheet, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Assign a SUM formula to cell B2 (example: sum of A1:A10)
        Cell targetCell = sheet.Cells["B2"];
        targetCell.Formula = "=SUM(A1:A10)";

        // Save the workbook with the updated formula
        workbook.Save("output.xlsx");
    }
}
