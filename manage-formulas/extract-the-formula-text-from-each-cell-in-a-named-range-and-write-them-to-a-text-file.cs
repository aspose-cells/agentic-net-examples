// Title: Extract cell formulas from a named range and save them to a text file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells extract formulas from a specific named range to a .txt file | How to write Excel cell formulas to a text file using Aspose.Cells .NET | Retrieve formula text for each cell in a named range with Aspose.Cells in C# | Save named range cell addresses and formulas to a plain text file using Aspose.Cells | Aspose.Cells iterate over named range cells and export formulas in C#
// Tags: named range formula extraction Aspose.Cells | export Excel formulas to text file C# | iterate cells in Aspose.Cells range | cell.IsFormula property Aspose.Cells | write cell address and formula to file C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, locates a named range, iterates through each cell in that range, captures the formula text (or the cell value if no formula), and writes the cell address together with the extracted formula to a specified text file, creating the output directory when necessary.
class ExtractFormulasFromNamedRange
{
    static void Main()
    {
        // Path to the Excel workbook
        string workbookPath = "input.xlsx";

        // Name of the named range to process
        string rangeName = "MyRange";

        // Path to the output text file
        string outputPath = "formulas.txt";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file \"{workbookPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the named range collection
            NameCollection names = workbook.Worksheets.Names;

            // Find the specific named range
            Name namedRange = names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Get the cells area of the named range (use fully qualified type to avoid ambiguity)
            Aspose.Cells.Range range = namedRange.GetRange();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Prepare a StreamWriter for the output file
            using (StreamWriter writer = new StreamWriter(outputPath, false))
            {
                // Iterate through each cell in the range
                foreach (Cell cell in range)
                {
                    // Retrieve the formula text; if the cell has no formula, use its value as a fallback
                    string formula = cell.IsFormula ? cell.Formula : cell.Value?.ToString() ?? string.Empty;

                    // Write the cell address and its formula/value to the file
                    writer.WriteLine($"{cell.Name}: {formula}");
                }
            }

            Console.WriteLine($"Formulas extracted to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
