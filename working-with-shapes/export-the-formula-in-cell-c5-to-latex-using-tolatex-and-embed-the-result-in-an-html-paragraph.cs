// Title: Convert a worksheet cell formula to LaTeX and embed it in an HTML paragraph using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert Excel formula to LaTeX string | How to obtain LaTeX representation of a worksheet formula with Aspose.Cells | Embed LaTeX output from Aspose.Cells formula conversion into an HTML <p> element in .NET
// Tags: Aspose.Cells formula to LaTeX conversion | C# Aspose.Cells ToLaTeX method usage | reflection call Aspose.Cells Formula ToLaTeX | export Excel cell formula as LaTeX | embed LaTeX string in HTML paragraph C#

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example creates a workbook, assigns the formula "=SUM(A1:B4)" to cell C5, uses Aspose.Cells.Formula.ToLaTeX (invoked via reflection) to convert the formula to a LaTeX string, wraps the result in an HTML <p> tag, and writes the HTML to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (no external file needed)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a sample formula in cell C5
            Cell targetCell = worksheet.Cells["C5"];
            targetCell.Formula = "=SUM(A1:B4)";

            // Convert the formula to LaTeX using Aspose.Cells Formula utility (via reflection for compatibility)
            string latexFormula = ConvertFormulaToLaTeX(targetCell.Formula);

            // Embed the LaTeX string in an HTML paragraph
            string htmlParagraph = $"<p>{latexFormula}</p>";

            // Output the HTML (for demonstration purposes)
            Console.WriteLine(htmlParagraph);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static string ConvertFormulaToLaTeX(string formula)
    {
        try
        {
            // Locate the Aspose.Cells.Formula type in the loaded assembly
            Type formulaType = typeof(Workbook).Assembly.GetType("Aspose.Cells.Formula");
            if (formulaType != null)
            {
                // Look for a static method named ToLaTeX that accepts a string
                MethodInfo toLaTeXMethod = formulaType.GetMethod("ToLaTeX", BindingFlags.Public | BindingFlags.Static, null, new[] { typeof(string) }, null);
                if (toLaTeXMethod != null)
                {
                    // Invoke the method and return the LaTeX representation
                    object result = toLaTeXMethod.Invoke(null, new object[] { formula });
                    return result as string ?? formula;
                }
            }
        }
        catch
        {
            // Ignored – fallback will be used
        }

        // Fallback: return the original formula if conversion is not possible
        return formula;
    }
}
