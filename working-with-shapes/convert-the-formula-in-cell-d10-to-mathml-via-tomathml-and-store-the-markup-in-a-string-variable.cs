// Title: C# – Convert an Excel cell formula to MathML using Aspose.Cells ToMathML()
// Description: This example creates a workbook, writes a formula to cell D10, places the formula in a TextBox shape, extracts the first equation paragraph, calls EquationNode.ToMathML() and stores the resulting MathML markup in a string variable.
// Keywords: Aspose.Cells | C# | .NET | MathML conversion | EquationNode.ToMathML | Excel formula to MathML | cell D10 | TextBox shape | equation paragraph | store MathML string
// Common Searches: Aspose.Cells convert formula to MathML C# | How to get MathML from Excel cell using Aspose | EquationNode ToMathML example | Retrieve MathML from TextBox shape Aspose.Cells | Convert SUM(A1:A5) to MathML .NET
// Developer Intent: Obtain a MathML string that represents the formula stored in cell D10.
// Use Cases: Embed MathML in web pages to display Excel formulas as scalable equations. | Save MathML markup in a database for reuse in scientific reports. | Generate PDFs that contain MathML‑derived equations from worksheet data.
// AI Prompts: Write C# code that reads the formula from cell D10 and returns its MathML using Aspose.Cells. | Explain the purpose of EquationNode.ToMathML and the MathML features it supports. | Show an alternative way to convert a cell formula to MathML without using a TextBox shape.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;

// This example creates a workbook, writes a formula to cell D10, places the formula in a TextBox shape, extracts the first equation paragraph, calls EquationNode.ToMathML() and stores the resulting MathML markup in a string variable.
class ConvertFormulaToMathML
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: set a formula in cell D10 (you can replace this with your own data)
        worksheet.Cells["D10"].Formula = "=SUM(A1:A5)";

        // Retrieve the formula string from cell D10
        string formula = worksheet.Cells["D10"].Formula;

        // Add a TextBox shape to hold the equation (position and size are arbitrary)
        TextBox textBox = worksheet.Shapes.AddTextBox(0, 0, 100, 20, 200, 50);

        // Set the TextBox text to the formula string.
        // Aspose.Cells will treat this as an equation paragraph.
        textBox.Text = formula;

        // Get the equation paragraph (first math paragraph) from the TextBox
        EquationNode equationNode = textBox.GetEquationParagraph();

        // Variable to store the MathML markup
        string mathML = string.Empty;

        if (equationNode != null)
        {
            // Convert the equation to MathML
            mathML = equationNode.ToMathML();

            // Output the MathML to console (optional)
            Console.WriteLine("MathML representation:");
            Console.WriteLine(mathML);
        }
        else
        {
            Console.WriteLine("No equation paragraph was created from the formula.");
        }

        // Save the workbook (optional, demonstrates create/save lifecycle)
        workbook.Save("FormulaToMathML.xlsx");
    }
}
