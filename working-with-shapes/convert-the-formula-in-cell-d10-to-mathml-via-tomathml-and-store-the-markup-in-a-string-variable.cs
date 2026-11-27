// Title: Convert an Excel cell formula to MathML using Aspose.Cells EquationNode in C#
// Common Searches: how to get MathML from a worksheet formula with Aspose.Cells C# | Aspose.Cells example converting cell D10 formula to MathML | C# create invisible equation shape to extract MathML using Aspose.Cells | retrieve MathML markup from Excel formula via EquationNode.ToMathML | store generated MathML in another worksheet cell with Aspose.Cells
// Tags: Aspose.Cells EquationNode ToMathML conversion | C# generate MathML from Excel formula | add invisible equation shape for MathML extraction | store MathML markup in worksheet cell | convert formula string to MathML using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing.Equations;

namespace AsposeCellsFormulaToMathML
{
    // The sample loads or creates a workbook containing a formula in D10, adds an invisible equation shape, obtains its EquationNodeParagraph, clears default children, inserts the formula as a TextRunEquationNode, calls ToMathML() on the root node to produce MathML markup stored in a string, writes the MathML to cell E10, auto‑fits columns, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains the formula in cell D10.
            // If the file does not exist, we create a sample workbook with a formula.
            string workbookPath = "input.xlsx";

            Workbook workbook;
            if (File.Exists(workbookPath))
            {
                // Load existing workbook
                workbook = new Workbook(workbookPath);
            }
            else
            {
                // Create a new workbook and add a sample formula to D10
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Example formula: =SUM(A1:A5)
                ws.Cells["D10"].Formula = "=SUM(A1:A5)";
                // Fill some sample data for the formula to be meaningful
                ws.Cells["A1"].PutValue(1);
                ws.Cells["A2"].PutValue(2);
                ws.Cells["A3"].PutValue(3);
                ws.Cells["A4"].PutValue(4);
                ws.Cells["A5"].PutValue(5);
                // Save the sample workbook for future runs
                workbook.Save(workbookPath);
                Console.WriteLine($"Sample workbook created at '{Path.GetFullPath(workbookPath)}'.");
            }

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the formula from cell D10
            string formula = worksheet.Cells["D10"].Formula;
            if (string.IsNullOrEmpty(formula))
            {
                Console.WriteLine("Cell D10 does not contain a formula.");
                return;
            }

            Console.WriteLine($"Formula in D10: {formula}");

            // ------------------------------------------------------------
            // Convert the formula to MathML using EquationNode.ToMathML()
            // ------------------------------------------------------------
            // Aspose.Cells does not provide a direct conversion from a formula string
            // to an EquationNode. The typical approach is to create an equation shape,
            // insert the formula text as a TextRunEquationNode, and then call ToMathML()
            // on the root EquationNode (which represents the whole equation paragraph).
            // This demo follows that approach.

            // 1. Add an invisible equation shape to the worksheet.
            //    Width and height are arbitrary because we only need the equation object.
            var equationShape = worksheet.Shapes.AddEquation(0, 0, 300, 100, 0, 0);

            // 2. Get the root EquationNodeParagraph (the container for the equation).
            var equationRoot = equationShape.GetEquationParagraph();

            // 3. Clear any default children (the shape may contain a default empty text node).
            equationRoot.RemoveAllChildren();

            // 4. Add a single TextRunEquationNode containing the formula string.
            //    Note: The formula string includes the leading '=', which is typical for Excel formulas.
            //    The EquationNode parser does not interpret the formula; it simply treats it as text.
            //    For a true mathematical representation, you would need to construct the equation
            //    using specific EquationNode types (Fraction, Superscript, etc.). Here we demonstrate
            //    the conversion mechanism using the raw text.
            var textNode = equationRoot.AddChild(EquationNodeType.Text) as TextRunEquationNode;
            textNode.Text = formula;

            // 5. Convert the equation to MathML.
            string mathML;
            try
            {
                mathML = equationRoot.ToMathML();
                Console.WriteLine("\nGenerated MathML:");
                Console.WriteLine(mathML);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting equation to MathML: {ex.Message}");
                return;
            }

            // Store the MathML markup in a string variable (already done as 'mathML')
            // For demonstration, write the MathML back to a cell (e.g., E10)
            worksheet.Cells["E10"].PutValue(mathML);

            // Optionally, auto‑fit the column to display the result nicely.
            worksheet.AutoFitColumns();

            // Save the workbook with the MathML result.
            string outputPath = "output_with_mathml.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved with MathML in cell E10: {Path.GetFullPath(outputPath)}");
        }
    }
}
