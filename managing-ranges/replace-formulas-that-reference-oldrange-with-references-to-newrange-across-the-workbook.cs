// Title: Replace a named range reference in every formula of an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# replace named range reference in all formulas of a workbook | programmatically rename a range in Excel formulas with Aspose.Cells .NET | update formula strings from OldRange to NewRange across worksheets using Aspose.Cells | bulk modify Excel formula references in C# via Aspose.Cells API | iterate through cells and change named range in formulas Aspose.Cells example
// Tags: named range formula replacement Aspose.Cells | bulk Excel formula update C# | worksheet cell iteration Aspose.Cells | programmatic range rename Excel | Aspose.Cells workbook formula manipulation

using Aspose.Cells;
using System;

// The sample loads an Excel workbook, walks through each worksheet and every cell, detects formulas, replaces any occurrence of the named range "OldRange" with "NewRange" in those formulas, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Replace references from "OldRange" to "NewRange"
                    string updatedFormula = cell.Formula.Replace("OldRange", "NewRange");
                    if (updatedFormula != cell.Formula)
                    {
                        cell.Formula = updatedFormula;
                    }
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
