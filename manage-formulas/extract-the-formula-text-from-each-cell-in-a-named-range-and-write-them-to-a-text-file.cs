// Title: C# – Export Formulas from a Named Range to a Text File with Aspose.Cells
// Description: Loads an Excel workbook, locates a defined name, iterates over every cell in the associated range(s), captures the formula text of formula cells, and writes each address‑formula pair to a plain‑text file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# extract formulas | named range formula export | write Excel formulas to txt | cell.IsFormula Aspose | retrieve defined name cells | .NET Excel formula extraction
// Common Searches: Aspose.Cells get formula text from a named range | export formulas to txt file C# | iterate cells in defined name Aspose | how to write Excel formulas to a text file
// Developer Intent: Read all formula strings inside a specific named range and save them as address‑formula lines in a text document.
// Use Cases: Create an audit trail of calculations in a financial model. | Generate documentation that lists custom formulas for review. | Compare exported formulas against a baseline to detect unexpected changes.
// AI Prompts: Generate C# code that extracts formulas from a named range and outputs them to a CSV file with Aspose.Cells. | Show how to include each cell's evaluated value together with its formula in the export. | Explain handling of multiple, non‑contiguous ranges returned by a named range when writing formulas to a file.

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel workbook, locates a defined name, iterates over every cell in the associated range(s), captures the formula text of formula cells, and writes each address‑formula pair to a plain‑text file using Aspose.Cells for .NET.
class ExtractFormulasFromNamedRange
{
    static void Main()
    {
        // Path to the source workbook
        string workbookPath = "input.xlsx";

        // Name of the defined range to process
        string rangeName = "MyRange";

        // Path for the output text file
        string outputPath = "formulas.txt";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file '{workbookPath}' not found.");
                return;
            }

            // Load the workbook from the file system
            Workbook workbook = new Workbook(workbookPath);

            // Retrieve the Name object that represents the named range
            Name namedRange = workbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range '{rangeName}' not found in the workbook.");
                return;
            }

            // Get all Range objects that the name refers to
            Aspose.Cells.Range[] ranges = namedRange.GetRanges();

            // Open a StreamWriter to write formulas to the text file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                foreach (Aspose.Cells.Range range in ranges)
                {
                    // Iterate through each cell in the current range
                    foreach (Cell cell in range)
                    {
                        // Write only cells that actually contain a formula
                        if (cell.IsFormula)
                        {
                            // Write the cell address (e.g., A1) followed by its formula
                            writer.WriteLine($"{cell.Name}: {cell.Formula}");
                        }
                    }
                }
            }

            Console.WriteLine($"Formulas from named range '{rangeName}' have been written to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
