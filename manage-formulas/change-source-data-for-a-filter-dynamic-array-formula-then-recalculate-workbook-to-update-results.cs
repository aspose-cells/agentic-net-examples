// Title: How to replace the source range of a FILTER dynamic array formula and recalculate the workbook using Aspose.Cells in C#
// Common Searches: aspnet replace source range in Excel FILTER formula using Aspose.Cells | c# update FILTER dynamic array source range and recalculate workbook | programmatically modify FILTER function range with Aspose.Cells .NET | recalculate Excel workbook after changing a formula with Aspose.Cells C# | Aspose.Cells change FILTER formula arguments and refresh results
// Tags: filter formula source range modification Aspose.Cells | workbook recalculation after formula edit .NET | dynamic array update using Aspose.Cells C# | replace formula substring Aspose.Cells API | excel filter function programmatic change

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, finds the FILTER dynamic array formula in cell A1, replaces its source range with "D2:D10", triggers a full workbook recalculation to update the array results, and saves the modified file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook (no password required)
            var loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Get the worksheet that contains the FILTER formula
            Worksheet sheet = workbook.Worksheets["Sheet1"];
            if (sheet == null)
            {
                Console.WriteLine("Worksheet 'Sheet1' not found.");
                return;
            }

            // Locate the cell with the FILTER dynamic array formula (example: A1)
            Cell filterCell = sheet.Cells["A1"];
            if (filterCell == null)
            {
                Console.WriteLine("Cell A1 not found.");
                return;
            }

            string originalFormula = filterCell.Formula;
            if (string.IsNullOrEmpty(originalFormula) ||
                originalFormula.IndexOf("FILTER(", StringComparison.OrdinalIgnoreCase) < 0)
            {
                Console.WriteLine("No FILTER formula found in cell A1.");
                return;
            }

            // Define the new source range you want the FILTER to use
            string newSourceRange = "D2:D10";

            // Replace the source range part of the FILTER formula
            int filterPos = originalFormula.IndexOf("FILTER(", StringComparison.OrdinalIgnoreCase);
            if (filterPos >= 0)
            {
                int openParen = originalFormula.IndexOf('(', filterPos);
                int firstComma = originalFormula.IndexOf(',', openParen);
                if (firstComma > openParen)
                {
                    string beforeSource = originalFormula.Substring(0, openParen + 1);
                    string afterSource = originalFormula.Substring(firstComma);
                    string newFormula = beforeSource + newSourceRange + afterSource;
                    filterCell.Formula = newFormula;
                }
            }

            // Recalculate the workbook so the dynamic array updates
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
