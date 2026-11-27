// Title: Convert an Excel cell formula to LaTeX and embed it in an HTML paragraph using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export cell C5 formula as LaTeX string | How to generate LaTeX from an Excel formula with Aspose.Cells .NET | Save LaTeX representation of a worksheet formula to an HTML file using Aspose.Cells | Aspose.Cells ToLaTeX example for converting a formula to an HTML paragraph | Retrieve Excel formula and convert to LaTeX in C# without persisting the shape
// Tags: Aspose.Cells ToLaTeX API usage | convert Excel formula to LaTeX in .NET | write LaTeX output to HTML paragraph C# | create transient equation shape for conversion | extract cell formula without saving workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing.Equations;

namespace AsposeCellsFormulaToLaTeX
{
    // The sample loads a workbook, accesses cell C5, obtains its formula (or displayed value if no formula), strips a leading '=', creates a temporary equation shape, inserts the formula as a text node, converts the equation node to a LaTeX string via ToLaTeX(), HTML‑encodes the result, and writes it inside a <p> element to an output HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input workbook (must exist)
            const string inputPath = "input.xlsx";
            // Path to the output HTML file
            const string outputPath = "output.html";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                // Use the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the cell C5
                Cell targetCell = worksheet.Cells["C5"];
                if (targetCell == null)
                {
                    Console.WriteLine("Error: Cell C5 does not exist.");
                    return;
                }

                // Get the formula string (if the cell contains a formula)
                string formula = targetCell.Formula;
                // If the cell does not contain a formula, fall back to its displayed value
                if (string.IsNullOrEmpty(formula))
                {
                    formula = targetCell.StringValue;
                    Console.WriteLine("Info: Cell C5 does not contain a formula. Using its displayed value instead.");
                }

                // Remove leading '=' if present – EquationNode expects plain text
                if (formula.StartsWith("="))
                {
                    formula = formula.Substring(1);
                }

                // Create a temporary equation shape (it will not be saved)
                // Parameters: upper left row, upper left column, width, height, upper left offset X, offset Y
                var shape = worksheet.Shapes.AddEquation(0, 0, 200, 50, 0, 0);

                // Get the root equation paragraph node
                EquationNode equationRoot = shape.GetEquationParagraph();

                // Add a text node containing the formula string
                var textNode = equationRoot.AddChild(EquationNodeType.Text) as TextRunEquationNode;
                if (textNode == null)
                {
                    Console.WriteLine("Error: Unable to create a text node for the equation.");
                    return;
                }
                textNode.Text = formula;

                // Convert the equation node to LaTeX
                string latexExpression = equationRoot.ToLaTeX();

                // Build the HTML paragraph
                string htmlParagraph = $"<p>{System.Net.WebUtility.HtmlEncode(latexExpression)}</p>";

                // Write the HTML to the output file
                File.WriteAllText(outputPath, htmlParagraph);
                Console.WriteLine("LaTeX conversion successful. HTML written to: " + Path.GetFullPath(outputPath));
                Console.WriteLine("Generated HTML:");
                Console.WriteLine(htmlParagraph);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An exception occurred: " + ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
