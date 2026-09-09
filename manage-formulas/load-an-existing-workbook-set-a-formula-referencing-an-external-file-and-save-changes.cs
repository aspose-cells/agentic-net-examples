// Title: Load an Excel workbook, add an external reference formula to a cell, recalculate, and save with Aspose.Cells for .NET (C#)
// AI Prompts: Insert a formula that links cell B2 to A1 in an external workbook (External.xlsx) and recalculate the workbook using Aspose.Cells in C#. | Load a .xlsx file, assign a cross‑file reference formula, trigger calculation, and write the updated workbook to a new file with Aspose.Cells.
// Common Searches: Aspose.Cells C# create formula that points to a cell in a different workbook | recalculate all formulas after adding a cross‑file link with Aspose.Cells | save modified workbook after changing cell formula using Aspose.Cells .NET | sample code for cross‑workbook cell reference in Aspose.Cells C# | load Excel file, add external reference, and write output.xlsx with Aspose.Cells
// Tags: add cross‑workbook formula Aspose.Cells | trigger formula calculation C# | open and save workbook Aspose.Cells .NET | set cell B2 external reference Aspose.Cells | programmatic Excel formula update C#

using System;
using Aspose.Cells;

// // Loads "input.xlsx", sets B2 to a formula referencing A1 in "External.xlsx", recalculates all formulas, and saves the result as "output.xlsx" using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula that references a cell in an external workbook.
        // Example: reference cell A1 in Sheet1 of External.xlsx
        sheet.Cells["B2"].Formula = "='[External.xlsx]Sheet1'!A1";

        // Recalculate formulas so the new value is evaluated (optional)
        workbook.CalculateFormula();

        // Save the modified workbook back to disk
        workbook.Save("output.xlsx");
    }
}
