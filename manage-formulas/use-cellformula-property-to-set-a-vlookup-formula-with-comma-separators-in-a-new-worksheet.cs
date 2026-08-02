// Title: Insert a VLOOKUP formula with commas using Cell.Formula in Aspose.Cells for .NET
// Description: Creates a workbook, builds a D‑E lookup table, sets the formula "=VLOOKUP(A2, D1:E4, 2, FALSE)" in B2 via Cell.Formula, forces calculation, and saves the file as VlookupDemo.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Cell.Formula | VLOOKUP | comma‑separated formula | calculate formulas | save workbook | Excel automation | lookup table
// Common Searches: Aspose.Cells set VLOOKUP formula with commas | Cell.Formula VLOOKUP example C# | how to calculate formulas after inserting VLOOKUP Aspose | save workbook after adding VLOOKUP Aspose.Cells | programmatically add lookup formula in Excel using Aspose
// Developer Intent: Programmatically add a comma‑separated VLOOKUP expression to a cell, evaluate it, and persist the workbook.
// Use Cases: Pre‑populate reports with VLOOKUP functions that auto‑calculate on open. | Generate data‑driven Excel files where lookup logic is embedded during creation. | Automate cross‑reference checks by inserting VLOOKUP formulas into source worksheets.
// AI Prompts: Write C# code with Aspose.Cells that adds a VLOOKUP formula using commas to cell B2, recalculates the workbook, and saves it as an .xlsx file. | Show how to use Cell.Formula to set "=VLOOKUP(A2, D1:E4, 2, FALSE)" and then call CalculateFormula in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, builds a D‑E lookup table, sets the formula "=VLOOKUP(A2, D1:E4, 2, FALSE)" in B2 via Cell.Formula, forces calculation, and saves the file as VlookupDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet (the first one is already present, we can use it)
        Worksheet worksheet = workbook.Worksheets[0];

        // ----- Prepare data for VLOOKUP -----
        // Header for lookup table
        worksheet.Cells["D1"].PutValue("Key");
        worksheet.Cells["E1"].PutValue("Value");

        // Sample lookup rows
        worksheet.Cells["D2"].PutValue("A");
        worksheet.Cells["E2"].PutValue(100);
        worksheet.Cells["D3"].PutValue("B");
        worksheet.Cells["E3"].PutValue(200);
        worksheet.Cells["D4"].PutValue("C");
        worksheet.Cells["E4"].PutValue(300);

        // Value we want to look up
        worksheet.Cells["A2"].PutValue("B");

        // ----- Set VLOOKUP formula using comma separators -----
        // Formula: =VLOOKUP(A2, D1:E4, 2, FALSE)
        worksheet.Cells["B2"].Formula = "=VLOOKUP(A2, D1:E4, 2, FALSE)";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("VlookupDemo.xlsx");
    }
}
