// Title: Extract formulas from a named range in an Excel workbook and write them to a text file with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, retrieves a specific named range, iterates over its cells, and outputs each cell's Formula string as a new line in a .txt file. | Generate a method that takes a workbook path, a named‑range identifier, and an output file path, then extracts all formulas from that range with Aspose.Cells and saves them line‑by‑line. | Create robust error‑handling logic for loading a workbook, locating a named range, and exporting its formulas to a text file, ensuring graceful messages when the file or range is missing.
// Common Searches: Aspose.Cells C# export formulas from a named range to a .txt file | how to retrieve formula strings of a named range using Aspose.Cells | save Excel named range formulas as plain text with Aspose.Cells .NET | C# extract formulas from specific named range in workbook | Aspose.Cells get formula text from range and write to file
// Tags: named range formula extraction Aspose.Cells | cell.Formula property usage C# | write Excel formulas to text file .NET | range traversal for formula export Aspose.Cells | error handling for named range access Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Loads input.xlsx, locates the named range "MyRange", iterates each cell in the range, reads the Cell.Formula value (empty if none), and writes each formula line‑by‑line to output.txt using Aspose.Cells for .NET.
class ExtractFormulasFromNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range by its name (replace "MyRange" with your actual range name)
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range \"MyRange\" not found.");
                return;
            }

            // Get the actual cell range that the name refers to
            Aspose.Cells.Range range = namedRange.GetRange();
            if (range == null)
            {
                Console.WriteLine("The named range does not refer to a valid cell range.");
                return;
            }

            // Write formulas to the output text file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                for (int row = 0; row < range.RowCount; row++)
                {
                    for (int col = 0; col < range.ColumnCount; col++)
                    {
                        Cell cell = range[row, col];
                        string formulaText = cell.Formula; // Empty string if no formula
                        writer.WriteLine(formulaText);
                    }
                }
            }

            Console.WriteLine($"Formulas have been extracted to \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
