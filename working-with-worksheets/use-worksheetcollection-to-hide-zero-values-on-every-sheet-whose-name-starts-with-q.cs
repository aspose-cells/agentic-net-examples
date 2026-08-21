// Title: Hide Zero Values on Worksheets Starting with “Q” Using Aspose.Cells (C#)
// Description: Loads an Excel workbook, iterates its WorksheetCollection, and disables the DisplayZeros property on every sheet whose name begins with "Q" (case‑insensitive). The modified workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | Hide zero values | DisplayZeros | WorksheetCollection | filter worksheets by name | Excel zero suppression | quarterly sheets | batch workbook processing
// Common Searches: Aspose.Cells hide zeros on specific worksheets | set DisplayZeros false for sheets starting with Q | C# iterate worksheets and suppress zero values | how to hide zero values in quarterly Excel sheets using Aspose | filter worksheets by prefix and change display settings
// Developer Intent: Programmatically hide zero values on every worksheet whose name starts with the letter "Q".
// Use Cases: Prepare quarterly reports (Q1, Q2, Q3, Q4) so that cells containing 0 are not displayed, improving readability for stakeholders. | Create a reusable Excel template where all Q‑prefixed sheets automatically suppress zero values before distribution. | Automate a nightly job that processes multiple workbooks, turning off zero display on every sheet whose name begins with "Q".
// AI Prompts: Generate C# code with Aspose.Cells that hides zero values on all worksheets whose names start with "Q" and saves the result as XLSX. | Explain how the DisplayZeros property works and how to apply it selectively to worksheets in a workbook. | Provide robust error‑handling patterns for loading a workbook, iterating its WorksheetCollection, and updating display settings with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, iterates its WorksheetCollection, and disables the DisplayZeros property on every sheet whose name begins with "Q" (case‑insensitive). The modified workbook is saved as an XLSX file.
    public class HideZeroValuesInQSheets
    {
        // Entry point for the console application
        public static void Main(string[] args)
        {
            try
            {
                // Expect input and output file paths as command‑line arguments
                if (args.Length < 2)
                {
                    Console.WriteLine("Usage: HideZeroValuesInQSheets <inputFilePath> <outputFilePath>");
                    return;
                }

                string inputFilePath = args[0];
                string outputFilePath = args[1];

                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputFilePath))
                    throw new FileNotFoundException($"Input file not found: {inputFilePath}");

                Run(inputFilePath, outputFilePath);
                Console.WriteLine($"Workbook processed and saved to: {outputFilePath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Core logic to hide zero values in worksheets whose names start with "Q"
        public static void Run(string inputFilePath, string outputFilePath)
        {
            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputFilePath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Hide zero values if the worksheet name starts with "Q"
                    if (sheet.Name.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet.DisplayZeros = false;
                    }
                }

                // Save the modified workbook to the desired output file
                workbook.Save(outputFilePath, SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                // Propagate errors to the caller
                throw new ApplicationException("Failed to process the workbook.", ex);
            }
        }
    }
}
