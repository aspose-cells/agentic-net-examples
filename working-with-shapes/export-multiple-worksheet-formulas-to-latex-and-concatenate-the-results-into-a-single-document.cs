// Title: Export formulas from every worksheet in an Excel workbook to a combined LaTeX document using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, iterates all worksheets, extracts each cell's formula, escapes LaTeX special characters, and writes the formulas with cell references into a single .tex file. | Demonstrate how to concatenate formulas from multiple sheets into one LaTeX file, adding worksheet comments and wrapping each formula in math mode, using the Aspose.Cells API.
// Common Searches: C# Aspose.Cells export all cell formulas to LaTeX file | How to write Excel formulas to a .tex document with Aspose.Cells | Iterate worksheets and extract formulas for LaTeX output in .NET | Escape underscores when converting Excel formulas to LaTeX using C# | Combine formulas from multiple Excel sheets into one LaTeX file
// Tags: export formulas to LaTeX Aspose.Cells | iterate worksheets Aspose.Cells C# | write .tex file from Excel formulas | LaTeX escaping for Excel formulas C# | concatenate sheet formulas into single LaTeX document

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// // Loads an Excel workbook, walks through each worksheet and cell containing a formula, escapes LaTeX characters, wraps formulas in math mode with cell references, and writes the aggregated output to a single .tex file.
class ExportFormulasToLatex
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // StringBuilder to accumulate LaTeX content
        StringBuilder latexBuilder = new StringBuilder();

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Add a comment indicating the worksheet name
            latexBuilder.AppendLine($"% Worksheet: {sheet.Name}");

            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Retrieve the formula string
                    string formula = cell.Formula;

                    // Basic LaTeX escaping (underscore is common in formulas)
                    formula = formula.Replace("_", "\\_");

                    // Append the formula wrapped in math mode, with cell reference for clarity
                    latexBuilder.AppendLine($"Cell {cell.Name}: ${formula}$");
                }
            }

            // Add a blank line between worksheets
            latexBuilder.AppendLine();
        }

        // Write the concatenated LaTeX to a .tex file
        File.WriteAllText("output.tex", latexBuilder.ToString());

        // Optional: inform the user that the operation completed
        Console.WriteLine("Formulas exported to LaTeX successfully.");
    }
}
