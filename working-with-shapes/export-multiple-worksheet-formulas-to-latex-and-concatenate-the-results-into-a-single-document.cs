// Title: Export Excel Formulas to LaTeX and Merge into One Sheet with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, scans all populated cells across worksheets, creates a temporary LaTeX equation shape for each formula using AddLaTeXEquation, converts the shape to a LaTeX string via EquationNode.ToLaTeX(), appends the cell reference and LaTeX code to a StringBuilder, writes the combined output to a new worksheet, and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | LaTeX export | Excel formula to LaTeX | AddLaTeXEquation | EquationNode ToLaTeX | concatenate LaTeX equations | multiple worksheets | export to new worksheet | save workbook with LaTeX
// Common Searches: Aspose.Cells export formulas to LaTeX C# | Convert Excel formulas to LaTeX with Aspose | Combine LaTeX equations from all sheets into one document | AddLaTeXEquation example Aspose.Cells | Save LaTeX output in a new worksheet
// Developer Intent: Extract every formula in a workbook, convert each to LaTeX, and concatenate the results into a single worksheet.
// Use Cases: Create a LaTeX‑based audit report that lists all calculations from a financial model. | Generate a single LaTeX file for scientific papers by pulling formulas from multiple experiment sheets. | Automate documentation of spreadsheet logic for compliance teams by exporting formulas as LaTeX.
// AI Prompts: Write C# code that uses Aspose.Cells to iterate over all worksheets, convert each cell formula to LaTeX with AddLaTeXEquation, and collect the results into one worksheet. | Show how to retrieve a LaTeX string from an EquationNode after adding a LaTeX equation shape. | Explain how to skip empty or non‑formula cells when exporting workbook formulas to LaTeX with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;
using System;
using System.Text;

// Loads an Excel file, scans all populated cells across worksheets, creates a temporary LaTeX equation shape for each formula using AddLaTeXEquation, converts the shape to a LaTeX string via EquationNode.ToLaTeX(), appends the cell reference and LaTeX code to a StringBuilder, writes the combined output to a new worksheet, and saves the workbook.
class ExportFormulasToLaTeX
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Collect LaTeX representations of all formulas
        StringBuilder latexBuilder = new StringBuilder();

        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Use the formula string as LaTeX input for a temporary equation shape
                    TextBox eqShape = worksheet.Shapes.AddLaTeXEquation(
                        topRow: row, top: 0,
                        leftColumn: col, left: 0,
                        height: 20, width: 200,
                        latex: cell.Formula);

                    // Retrieve the equation node and convert it to LaTeX
                    EquationNode eqNode = eqShape.GetEquationParagraph();
                    string latex = eqNode.ToLaTeX();

                    // Append the cell reference and its LaTeX expression
                    latexBuilder.AppendLine($"{cell.Name}: {latex}");
                }
            }
        }

        // Create a new worksheet to store the concatenated LaTeX output
        Worksheet resultSheet = workbook.Worksheets.Add("LaTeXExport");
        resultSheet.Cells[0, 0].PutValue(latexBuilder.ToString());

        // Save the workbook with the exported LaTeX content
        workbook.Save("output_with_latex.xlsx");
    }
}
