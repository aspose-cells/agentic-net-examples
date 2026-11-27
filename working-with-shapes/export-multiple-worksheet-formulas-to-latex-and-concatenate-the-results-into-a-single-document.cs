// Title: Export formulas from all Excel worksheets to a single LaTeX file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export Excel formulas to LaTeX document across multiple sheets | how to generate a combined .tex file with formulas from each worksheet using Aspose.Cells | iterate through workbook worksheets and write cell formulas as LaTeX in C# | Aspose.Cells example to create LaTeX align environment for Excel formulas | concatenate worksheet formula exports into one LaTeX file with Aspose.Cells
// Tags: export formulas to LaTeX with Aspose.Cells | C# extract worksheet formulas Aspose.Cells | generate LaTeX align environment from Excel formulas | concatenate multi‑sheet formula exports into single .tex | wrap Excel formula strings in LaTeX delimiters C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace ExportFormulasToLaTeX
{
    // The sample loads an Excel workbook with Aspose.Cells, iterates over every worksheet, extracts cells that contain formulas, wraps each formula in LaTeX math delimiters, and builds a LaTeX document with a subsection for each sheet. The resulting .tex file consolidates all formulas into a single document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Determine input file
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" does not exist.");
                    return;
                }

                // Load workbook inside a safe block
                Workbook workbook;
                try
                {
                    workbook = new Workbook(inputPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                    return;
                }

                var latexBuilder = new StringBuilder();

                // Optional LaTeX document header
                latexBuilder.AppendLine(@"\documentclass{article}");
                latexBuilder.AppendLine(@"\usepackage{amsmath}");
                latexBuilder.AppendLine(@"\begin{document}");
                latexBuilder.AppendLine(@"\section*{Exported Formulas}");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool sheetHasFormulas = false;
                    var sheetLatex = new StringBuilder();

                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;
                    int maxCol = cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = cells[row, col];
                            if (cell != null && cell.IsFormula)
                            {
                                // Retrieve the formula string (e.g., "=SUM(A1:A5)")
                                string formula = cell.Formula;

                                // Simple conversion: wrap the formula in LaTeX math delimiters.
                                string latexFormula = $"${formula}$";

                                // Build a line showing cell address and its LaTeX representation
                                sheetLatex.AppendLine($"{cell.Name}: {latexFormula}\\\\");
                                sheetHasFormulas = true;
                            }
                        }
                    }

                    if (sheetHasFormulas)
                    {
                        latexBuilder.AppendLine($@"\subsection*{{Worksheet: {sheet.Name}}}");
                        latexBuilder.AppendLine(@"\begin{align*}");
                        latexBuilder.Append(sheetLatex.ToString());
                        latexBuilder.AppendLine(@"\end{align*}");
                    }
                }

                // Close LaTeX document
                latexBuilder.AppendLine(@"\end{document}");

                // Write to output file
                string outputDir = Path.GetDirectoryName(inputPath) ?? ".";
                string outputPath = Path.Combine(outputDir, "Formulas.tex");
                try
                {
                    File.WriteAllText(outputPath, latexBuilder.ToString(), Encoding.UTF8);
                    Console.WriteLine($"LaTeX document generated successfully at \"{outputPath}\".");
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write LaTeX file: {writeEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
