// Title: C# Aspose.Cells: Convert formulas in column H to MathML and write to column I
// Description: A .NET console app that opens an Excel workbook, scans column H for formula cells, safely escapes each formula, wraps it in <math> tags to produce simple MathML, stores the markup in the adjacent column I, and saves the updated file. Ideal for automating spreadsheet‑to‑web content pipelines.
// Keywords: Aspose.Cells | C# | .NET | MathML generation | Excel formula conversion | column H to column I | XML escape | spreadsheet automation | batch processing | global
// Common Searches: convert Excel formulas to MathML C# Aspose.Cells | write MathML to adjacent cell in Excel using .NET | Aspose.Cells generate MathML from column H | C# code to export formulas as MathML | how to add MathML markup to Excel workbook
// Developer Intent: Create MathML for every formula in column H and place the markup in column I of the same worksheet.
// Use Cases: Embed MathML alongside scientific formulas for web publishing. | Prepare spreadsheets for HTML or PDF export with native MathML support. | Run a nightly job that enriches multiple workbooks with MathML for downstream processing.
// AI Prompts: Generate C# Aspose.Cells code that reads formulas from column H, escapes them, wraps them in <math> tags, and writes the result to column I with robust error handling. | Suggest improvements to include <mrow> and other MathML elements while preserving the existing loop logic. | Refactor the sample to log rows lacking formulas and skip empty cells efficiently.

using System;
using System.IO;
using Aspose.Cells;

namespace MathMLGenerator
{
    // A .NET console app that opens an Excel workbook, scans column H for formula cells, safely escapes each formula, wraps it in <math> tags to produce simple MathML, stores the markup in the adjacent column I, and saves the updated file. Ideal for automating spreadsheet‑to‑web content pipelines.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the last used row in the worksheet
                int lastRow = worksheet.Cells.MaxDataRow;

                // Iterate through each cell in column H (zero‑based index 7)
                for (int row = 0; row <= lastRow; row++)
                {
                    Cell formulaCell = worksheet.Cells[row, 7]; // Column H

                    // Process only cells that contain a formula
                    if (!string.IsNullOrEmpty(formulaCell.Formula))
                    {
                        try
                        {
                            // Simple MathML generation: wrap the formula in <math> tags.
                            // Escape special XML characters to ensure well‑formed output.
                            string escapedFormula = System.Security.SecurityElement.Escape(formulaCell.Formula);
                            string mathML = $"<math>{escapedFormula}</math>";

                            // Store the MathML markup in the adjacent cell (column I, index 8)
                            worksheet.Cells[row, 8].PutValue(mathML);
                        }
                        catch (Exception exCell)
                        {
                            Console.WriteLine($"Error processing formula at row {row + 1}: {exCell.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Processing completed. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
