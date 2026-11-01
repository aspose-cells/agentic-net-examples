// Title: Export formulas from all worksheets to a single LaTeX document with Aspose.Cells for .NET (C#)
// Common Searches: c# export all worksheet formulas to a .tex file using Aspose.Cells | how to convert Excel formulas to LaTeX align environment in a .NET application | generate combined LaTeX document from multiple Excel sheets with Aspose.Cells | Aspose.Cells iterate through worksheets and write formulas as LaTeX equations
// Tags: Aspose.Cells export formulas to LaTeX | C# create LaTeX document from Excel data | iterate Excel worksheets extract formulas Aspose.Cells | escape special characters for LaTeX in C# | replace Excel operators with LaTeX syntax

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, extracts cells that contain formulas, converts each formula to a simple LaTeX representation while escaping special characters and mapping common Excel operators, and appends the results as aligned equations inside a LaTeX document. All worksheet sections are concatenated and written to a single .tex file.
class ExportFormulasToLatex
{
    static void Main()
    {
        // Load the workbook from a file (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // StringBuilder to accumulate LaTeX content
        StringBuilder latexBuilder = new StringBuilder();

        // Optional: add LaTeX document preamble
        latexBuilder.AppendLine(@"\documentclass{article}");
        latexBuilder.AppendLine(@"\usepackage{amsmath}");
        latexBuilder.AppendLine(@"\begin{document}");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Add a section header for the worksheet
            latexBuilder.AppendLine($@"\section*{{{EscapeLatex(sheet.Name)}}}");
            latexBuilder.AppendLine(@"\begin{align*}");

            // Iterate over all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.IsFormula)
                {
                    // Get the cell address (e.g., A1) and the formula string
                    string address = cell.Name;
                    string formula = cell.Formula;

                    // Simple conversion of the formula to LaTeX:
                    // - Escape backslashes and special LaTeX characters
                    // - Wrap the formula in math mode
                    string latexFormula = ConvertFormulaToLatex(formula);

                    // Append the formula as an aligned equation: address = formula
                    latexBuilder.AppendLine($@"{EscapeLatex(address)} &= {latexFormula} \\");
                }
            }

            latexBuilder.AppendLine(@"\end{align*}");
            latexBuilder.AppendLine(); // Add a blank line between worksheets
        }

        // End of LaTeX document
        latexBuilder.AppendLine(@"\end{document}");

        // Write the concatenated LaTeX content to a .tex file
        File.WriteAllText("FormulasOutput.tex", latexBuilder.ToString(), Encoding.UTF8);
    }

    // Helper method to escape LaTeX special characters in plain text
    private static string EscapeLatex(string text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;

        return text
            .Replace(@"\", @"\textbackslash ")
            .Replace("{", @"\{")
            .Replace("}", @"\}")
            .Replace("#", @"\#")
            .Replace("$", @"\$")
            .Replace("%", @"\%")
            .Replace("&", @"\&")
            .Replace("^", @"\^{}")
            .Replace("_", @"\_")
            .Replace("~", @"\~{}");
    }

    // Very basic conversion of an Excel formula to LaTeX syntax
    private static string ConvertFormulaToLatex(string formula)
    {
        if (string.IsNullOrEmpty(formula))
            return string.Empty;

        // Remove leading '=' if present
        if (formula.StartsWith("="))
            formula = formula.Substring(1);

        // Escape LaTeX special characters inside the formula
        string escaped = EscapeLatex(formula);

        // Replace common Excel operators with LaTeX equivalents
        escaped = escaped
            .Replace("^", @"\hat{}")          // exponentiation (very simplistic)
            .Replace("*", @"\times ")
            .Replace("/", @"\div ")
            .Replace("<>", @"\neq ")
            .Replace("<=", @"\leq ")
            .Replace(">=", @"\geq ");

        // Wrap the result in math mode
        return "$" + escaped + "$";
    }
}
