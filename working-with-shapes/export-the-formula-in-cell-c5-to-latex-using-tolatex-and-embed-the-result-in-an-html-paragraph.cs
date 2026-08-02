// Title: Export an Excel cell formula to LaTeX and embed it in an HTML paragraph with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, writes the formula "=SUM(A1:B1)" to cell C5, adds an equation shape at the same location, copies the formula text into the shape, converts it to LaTeX using ToLaTeX(), wraps the result in a <p> tag, and saves the HTML paragraph to a file. The workbook is also saved for reference.
// Keywords: Aspose.Cells | C# | .NET | ToLaTeX | Excel formula to LaTeX | export formula as HTML | equation shape | cell C5 | SUM formula | HTML paragraph | LaTeX conversion | Excel automation
// Common Searches: convert Excel formula to LaTeX C# Aspose.Cells | export cell formula as HTML paragraph | ToLaTeX example for .NET | how to add equation shape in Aspose.Cells | save LaTeX output to HTML file
// Developer Intent: Generate a LaTeX string from a worksheet formula and write it inside an HTML <p> element using Aspose.Cells for .NET.
// Use Cases: Build web reports that display Excel calculations as rendered LaTeX equations. | Automate documentation pipelines that extract formulas from workbooks and embed them in HTML pages. | Create printable HTML content with LaTeX‑formatted equations for scientific or financial publications.
// AI Prompts: Show how to iterate over a range of cells and write each formula as a separate LaTeX paragraph in one HTML file. | Provide code that saves the LaTeX output to a .tex file instead of embedding it in HTML. | Explain how to adjust the size and position of the equation shape before calling ToLaTeX() for optimal LaTeX rendering.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;

// This example creates a workbook, writes the formula "=SUM(A1:B1)" to cell C5, adds an equation shape at the same location, copies the formula text into the shape, converts it to LaTeX using ToLaTeX(), wraps the result in a <p> tag, and saves the HTML paragraph to a file. The workbook is also saved for reference.
class ExportFormulaToLaTeX
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a sample formula in cell C5
        worksheet.Cells["C5"].Formula = "=SUM(A1:B1)";

        // Add an equation shape near cell C5 (parameters: topRow, top, leftColumn, left, height, width)
        TextBox equationShape = worksheet.Shapes.AddEquation(worksheet.Cells["C5"].Row, 0,
                                                            worksheet.Cells["C5"].Column, 0,
                                                            100, 300);

        // Retrieve the equation node from the shape
        EquationNode equationNode = equationShape.GetEquationParagraph();

        // Assign the cell formula text to the equation shape (as plain text)
        // This allows the ToLaTeX method to generate a LaTeX representation of the text.
        equationShape.Text = worksheet.Cells["C5"].Formula;

        // Convert the equation node to LaTeX
        string latexExpression = equationNode.ToLaTeX();

        // Embed the LaTeX expression in an HTML paragraph
        string htmlParagraph = $"<p>{latexExpression}</p>";

        // Write the HTML paragraph to a file
        File.WriteAllText("FormulaLaTeX.html", htmlParagraph);

        // Optionally, save the workbook (not required for the HTML output)
        workbook.Save("ExportFormulaToLaTeX.xlsx");
    }
}
