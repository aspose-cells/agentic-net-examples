using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;

class ExportFormulaToLaTeX
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Place a sample formula in cell C5
        worksheet.Cells["C5"].Formula = "=SUM(A1:B1)";

        // Add a textbox shape to the worksheet; this will host the equation
        // Parameters: topRow, top, leftColumn, left, height, width
        TextBox textBox = worksheet.Shapes.AddTextBox(5, 0, 2, 0, 30, 200);

        // Set the textbox text to the cell formula (treated as OfficeMath)
        textBox.Text = worksheet.Cells["C5"].Formula;

        // Retrieve the equation paragraph from the textbox
        EquationNode equationNode = textBox.GetEquationParagraph();

        // Convert the equation to LaTeX; if no equation node is found, fall back to the raw formula string
        string latexExpression = equationNode != null ? equationNode.ToLaTeX() : worksheet.Cells["C5"].Formula;

        // Embed the LaTeX expression inside an HTML paragraph
        string htmlContent = $"<p>{latexExpression}</p>";

        // Write the HTML to a file
        File.WriteAllText("FormulaLaTeX.html", htmlContent);

        // Optionally, save the workbook as HTML (preserving formulas)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportFormula = true
        };
        workbook.Save("Workbook.html", htmlOptions);
    }
}