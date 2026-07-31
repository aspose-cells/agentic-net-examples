// Title: Aspose.Cells for .NET – Replace “N/A” Formula Results with Blank Cells in a Named Range
// Description: Load an Excel workbook, locate a named range, calculate all formulas, and clear any cell whose evaluated value is the text "N/A". The modified workbook is saved as a new file, eliminating placeholder text from the specified range.
// Keywords: Aspose.Cells C# replace N/A | clear N/A cells named range | blank cells formula result Aspose | .NET Excel clean up N/A | Aspose.Cells range iteration
// Common Searches: Aspose.Cells replace N/A with blank in named range | C# clear cells that show N/A after calculation | How to remove N/A text from a specific range using Aspose.Cells | Excel formula result N/A to empty cell .NET
// Developer Intent: Remove cells that display the string "N/A" after formula evaluation within a defined named range.
// Use Cases: Prepare a client‑ready report by erasing placeholder N/A values from a designated area. | Clean data before exporting to analytics tools, ensuring no literal "N/A" strings remain. | Standardize financial models so cells that evaluate to N/A appear empty in the final workbook.
// AI Prompts: Write C# code with Aspose.Cells that clears cells showing "N/A" inside a named range. | Explain how to extend the loop to also clear cells that return the error #N/A while preserving other results. | Show an alternative approach using Aspose.Cells range operations to replace "N/A" text with blanks.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an Excel workbook, locate a named range, calculate all formulas, and clear any cell whose evaluated value is the text "N/A". The modified workbook is saved as a new file, eliminating placeholder text from the specified range.
    public class ReplaceNaFormulasWithBlank
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Name of the range to process
            const string rangeName = "MyRange";

            // Retrieve the named range object
            Name namedRange = workbook.Worksheets.Names[rangeName];
            if (namedRange == null)
            {
                Console.WriteLine($"Named range \"{rangeName}\" not found.");
                return;
            }

            // Get the address the name refers to (e.g., "=Sheet1!$A$1:$C$10")
            string refersTo = namedRange.RefersTo;
            if (string.IsNullOrEmpty(refersTo))
            {
                Console.WriteLine($"Named range \"{rangeName}\" does not have a valid reference.");
                return;
            }

            // Remove leading '=' if present
            string address = refersTo.TrimStart('=');

            // Split worksheet name and cell range (e.g., "Sheet1!$A$1:$C$10")
            int exclPos = address.IndexOf('!');
            if (exclPos < 0)
            {
                Console.WriteLine($"Invalid reference format for named range \"{rangeName}\".");
                return;
            }

            string sheetName = address.Substring(0, exclPos);
            string cellRange = address.Substring(exclPos + 1);

            // Access the worksheet
            Worksheet worksheet = workbook.Worksheets[sheetName];
            if (worksheet == null)
            {
                Console.WriteLine($"Worksheet \"{sheetName}\" not found.");
                return;
            }

            // Ensure all formulas are calculated before inspection
            workbook.CalculateFormula();

            // Create a Range object for the specified address (fully qualified to avoid ambiguity)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange(cellRange);

            // Iterate through each cell in the range
            for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
            {
                for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
                {
                    Cell cell = worksheet.Cells[row, col];

                    // Process only cells that contain a formula
                    if (cell.IsFormula)
                    {
                        // After calculation, check if the displayed value is the text "N/A"
                        if (cell.StringValue == "N/A")
                        {
                            // Replace the cell content with a blank (clear the cell)
                            cell.PutValue(string.Empty);
                        }
                    }
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Processing complete. Workbook saved as \"{outputPath}\".");
        }
    }
}
