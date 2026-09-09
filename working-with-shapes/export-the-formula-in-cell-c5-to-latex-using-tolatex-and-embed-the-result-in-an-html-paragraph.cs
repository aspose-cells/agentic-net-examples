// Title: Convert the formula in cell C5 to LaTeX with Aspose.Cells and write it into an HTML <p> element using C#
// AI Prompts: Generate C# code that loads an Excel workbook, reads the formula from cell C5, invokes Aspose.Cells.FormulaParser.ToLaTeX via reflection, wraps the LaTeX string in <p> tags, and saves the result to an HTML file. | Create a helper method that safely calls Aspose.Cells.FormulaParser.ToLaTeX through reflection and returns the original formula when the ToLaTeX method is not available.
// Common Searches: how to export Excel cell formula as LaTeX using Aspose.Cells C# | Aspose.Cells FormulaParser ToLaTeX example with reflection | write LaTeX output of Excel formula to HTML file in .NET | fallback handling when Aspose.Cells ToLaTeX method is missing | convert Excel formula to LaTeX and embed in HTML paragraph C#
// Tags: Aspose.Cells LaTeX conversion with reflection | C# write LaTeX output to HTML paragraph | Excel formula LaTeX export with Aspose.Cells | fallback handling for missing ToLaTeX method | cell C5 formula extraction Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Reflection;

// The example loads an Excel workbook, extracts the formula from cell C5, attempts to convert it to LaTeX using Aspose.Cells.FormulaParser.ToLaTeX via reflection, falls back to the original formula if the conversion API is unavailable, wraps the result in an HTML <p> element, and writes the HTML to a file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access cell C5
            Cell cell = worksheet.Cells["C5"];

            // Retrieve the formula (without the leading '=')
            string formula = cell.Formula;

            // Convert the formula to LaTeX using Aspose.Cells.FormulaParser if available
            string latex = ConvertFormulaToLaTeX(formula);

            // Embed the LaTeX string inside an HTML paragraph
            string htmlParagraph = $"<p>{latex}</p>";

            // Write the HTML content to a file
            File.WriteAllText(outputPath, htmlParagraph);

            Console.WriteLine($"LaTeX representation written to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Attempts to use Aspose.Cells.FormulaParser.ToLaTeX via reflection.
    // Falls back to returning the original formula if the API is unavailable.
    private static string ConvertFormulaToLaTeX(string formula)
    {
        try
        {
            // Look for the FormulaParser type in the Aspose.Cells assembly
            Type parserType = Type.GetType("Aspose.Cells.FormulaParser, Aspose.Cells");
            if (parserType != null)
            {
                MethodInfo toLatexMethod = parserType.GetMethod(
                    "ToLaTeX",
                    BindingFlags.Public | BindingFlags.Static,
                    null,
                    new[] { typeof(string) },
                    null);

                if (toLatexMethod != null)
                {
                    object result = toLatexMethod.Invoke(null, new object[] { formula });
                    return result as string ?? formula;
                }
            }
        }
        catch
        {
            // Swallow any reflection errors; fallback will be used.
        }

        // Fallback: return the original formula if conversion is not possible
        return formula;
    }
}
