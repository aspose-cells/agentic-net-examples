// Title: Convert Excel Cell Formula to LaTeX and Embed in HTML with Aspose.Cells (C#)
// Description: Shows how to read the formula in cell C5, place it in a textbox shape, obtain its EquationNode, convert the node to a LaTeX string via ToLaTeX(), wrap the string in an HTML <p> element, and save both the HTML file and the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | ToLaTeX | Excel formula to LaTeX | EquationNode | textbox shape | HTML export | LaTeX conversion | cell C5 | programmatic Excel | document automation
// Common Searches: Aspose.Cells convert formula to LaTeX C# | ToLaTeX example with EquationNode | Export Excel formula as LaTeX HTML | Create equation node from textbox shape Aspose.Cells | Save LaTeX string to HTML file using Aspose
// Developer Intent: Extract the formula from cell C5, transform it into LaTeX, and embed the result in an HTML paragraph.
// Use Cases: Generate LaTeX snippets for technical documentation directly from Excel calculations. | Produce web‑ready reports that display spreadsheet formulas as formatted equations. | Automate conversion of Excel logic into LaTeX for scientific publishing or e‑learning platforms.
// AI Prompts: Write C# code that reads a formula from a given cell, converts it to LaTeX with Aspose.Cells, and writes the output into an HTML file. | Explain the role of EquationNode and the ToLaTeX() method when converting a textbox shape equation. | Suggest ways to obtain an EquationNode without adding a temporary textbox shape.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;

namespace AsposeCellsFormulaToLaTeX
{
    // Shows how to read the formula in cell C5, place it in a textbox shape, obtain its EquationNode, convert the node to a LaTeX string via ToLaTeX(), wrap the string in an HTML <p> element, and save both the HTML file and the workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a sample formula into cell C5
            worksheet.Cells["C5"].Formula = "=A1+B1";

            // Retrieve the formula string from cell C5
            string excelFormula = worksheet.Cells["C5"].Formula; // e.g., "=A1+B1"

            // Add a temporary textbox shape to hold the formula as an equation
            // Parameters: topRow, top, leftColumn, left, height, width
            TextBox textBox = worksheet.Shapes.AddTextBox(4, 0, 4, 0, 200, 50);

            // Set the textbox text to the Excel formula.
            // Aspose will treat this as an equation paragraph.
            textBox.Text = excelFormula;

            // Get the equation paragraph (EquationNode) from the textbox
            EquationNode equationNode = textBox.GetEquationParagraph();

            // Convert the equation node to LaTeX using ToLaTeX()
            string latexExpression = equationNode != null ? equationNode.ToLaTeX() : string.Empty;

            // Embed the LaTeX expression in an HTML paragraph
            string htmlContent = $"<p>{latexExpression}</p>";

            // Save the HTML content to a file
            File.WriteAllText("FormulaLaTeX.html", htmlContent);

            // Optionally, save the workbook to verify the setup (lifecycle: save)
            workbook.Save("FormulaDemo.xlsx");
        }
    }
}
