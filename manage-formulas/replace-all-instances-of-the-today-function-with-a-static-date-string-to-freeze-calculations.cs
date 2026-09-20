// Title: Freeze TODAY() formulas by substituting a static date string in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that loads a workbook, walks through every worksheet and cell, and replaces each TODAY() call in a formula with a quoted fixed date string, then saves the result. | Generate a method that accepts an input file path, a static date value, and an output path, and updates all formulas containing TODAY() to use the provided date while preserving other parts of the expression. | Create a snippet that detects cells whose formula is exactly TODAY() and converts them to a plain date value, handling TODAY() embedded in larger expressions as well.
// Common Searches: C# Aspose.Cells replace TODAY() function with a constant date in all worksheets | How to freeze dynamic TODAY() values in an Excel file using Aspose.Cells .NET | Iterate through cells and substitute TODAY() in formulas with a fixed date string Aspose.Cells | Save workbook after converting TODAY() formulas to static dates in C# | Example of changing TODAY() to a hard‑coded date in Excel using Aspose.Cells API
// Tags: replace TODAY() with constant date Aspose.Cells | freeze dynamic date values in Excel Aspose.Cells | modify formulas across worksheets C# | convert TODAY() formula to static value .NET | save workbook after formula update Aspose.Cells | static date string substitution Excel formulas

using System;
using Aspose.Cells;

// The program loads an Excel file, iterates through every worksheet and cell, replaces any TODAY() occurrences in formulas with a fixed date string (as a value for standalone TODAY() formulas or as a quoted string within larger formulas), and saves the modified workbook to a new file.
class FreezeTodayFunction
{
    static void Main()
    {
        // Load the existing workbook (lifecycle rule: load)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Define the static date string to replace TODAY()
        string staticDate = DateTime.Today.ToString("yyyy-MM-dd");

        // Iterate through all worksheets and cells
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Use the Cells iterator for efficient traversal
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (cell.IsFormula && cell.Formula != null && cell.Formula.IndexOf("TODAY()", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    // If the formula is exactly TODAY(), replace the whole cell with the static date value
                    if (cell.Formula.Trim().Equals("TODAY()", StringComparison.OrdinalIgnoreCase))
                    {
                        cell.PutValue(staticDate);
                    }
                    else
                    {
                        // Replace each occurrence of TODAY() within the formula with a quoted static date string
                        string newFormula = cell.Formula.Replace("TODAY()", $"\"{staticDate}\"", StringComparison.OrdinalIgnoreCase);
                        cell.Formula = newFormula;
                    }
                }
            }
        }

        // Save the modified workbook (lifecycle rule: save)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
