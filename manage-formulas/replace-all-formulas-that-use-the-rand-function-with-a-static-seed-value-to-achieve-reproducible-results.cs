// Title: How to replace RAND() formulas with a fixed seed value in an Excel file using Aspose.Cells for .NET
// AI Prompts: Scan all worksheets in a workbook and substitute any RAND() formula with a predefined constant using Aspose.Cells C#. | Replace volatile RAND() functions with a static numeric seed to make random results reproducible in an Excel file via Aspose.Cells. | Programmatically convert cells containing RAND() to a fixed value and save the updated workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells replace RAND() with constant value in C# | make Excel RAND function deterministic using Aspose.Cells .NET | C# code to iterate over cells and change volatile formulas to static numbers | how to set a fixed value for random numbers in an Excel workbook with Aspose.Cells | convert RAND() formulas to fixed numbers across all sheets Aspose.Cells
// Tags: substitute RAND() with constant Aspose.Cells | static seed for Excel random function | iterate worksheets cells Aspose.Cells C# | convert volatile formulas to constants Aspose.Cells | deterministic random numbers Excel .NET

using Aspose.Cells;
using System;

// // Loads an Excel workbook, walks through every worksheet and cell, replaces each RAND() formula with a predefined static seed (e.g., 0.5), and saves the modified file to a new location.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Define a static seed value to replace RAND() formulas
        double staticSeed = 0.5; // you can set any reproducible value here

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the used range of cells in the current worksheet
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Loop through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a formula that uses RAND()
                    if (cell.IsFormula && cell.Formula.IndexOf("RAND", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the RAND() formula with the static seed value
                        cell.PutValue(staticSeed);
                    }
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
