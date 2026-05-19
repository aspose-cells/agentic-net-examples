using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Equations;

namespace AsposeCellsFormulaToMathML
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a sample formula in cell D10
            worksheet.Cells["D10"].Formula = "=SUM(A1:A5)";

            // Retrieve the formula text from the cell (including the leading '=')
            string formulaText = worksheet.Cells["D10"].Formula;

            // Add a textbox shape to the worksheet and assign the formula text to it.
            // This allows us to obtain an EquationNode which provides the ToMathML() method.
            TextBox textBox = worksheet.Shapes.AddTextBox(
                topRow: 0, top: 0, leftColumn: 0, left: 0,
                height: 100, width: 300);
            textBox.Text = formulaText;

            // Get the first equation paragraph (EquationNode) from the textbox
            EquationNode equationNode = textBox.GetEquationParagraph();

            // Convert the equation to MathML and store it in a string variable
            string mathML = string.Empty;
            if (equationNode != null)
            {
                mathML = equationNode.ToMathML();
            }

            // Output the MathML string (optional)
            Console.WriteLine("MathML representation of the formula in D10:");
            Console.WriteLine(mathML);

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("FormulaToMathML.xlsx");
        }
    }
}