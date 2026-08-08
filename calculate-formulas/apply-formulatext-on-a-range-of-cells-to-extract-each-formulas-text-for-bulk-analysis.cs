// Title: C# – Extract Formula Text from a Cell Range with Aspose.Cells and Save to a New Worksheet
// Description: Loads an Excel file (or creates a new workbook), defines a range (e.g., A1:C10), iterates each cell to obtain its non‑localized formula via GetFormula, records the address and formula string, writes the list to a new sheet named "FormulaAnalysis", and saves the workbook for bulk analysis.
// Keywords: Aspose.Cells C# formula extraction | GetFormula range | extract formulas from Excel | bulk formula analysis .NET | save formula list worksheet | Aspose.Cells example | Excel formula text retrieval
// Common Searches: Aspose.Cells get formula text from range | C# extract all formulas in A1:C10 | write extracted formulas to new sheet Aspose.Cells | how to list cell formulas in .NET | bulk formula extraction example
// Developer Intent: Retrieve the formula string of every cell in a specified range and export the address‑formula pairs to a separate worksheet for review or further processing.
// Use Cases: Audit a financial model by listing every formula in a defined area. | Generate documentation that shows spreadsheet logic without opening the file. | Feed extracted formulas into a validation engine that checks for prohibited functions.
// AI Prompts: Generate C# code that extracts formulas from a user‑defined range and writes them to a CSV file using Aspose.Cells. | Show how to include both the formula text and the evaluated value for each cell in the analysis sheet. | Explain how to obtain localized formula strings and compare them with the non‑localized version in Aspose.Cells.

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

namespace AsposeCellsExamples
{
    // Loads an Excel file (or creates a new workbook), defines a range (e.g., A1:C10), iterates each cell to obtain its non‑localized formula via GetFormula, records the address and formula string, writes the list to a new sheet named "FormulaAnalysis", and saves the workbook for bulk analysis.
    class FormulaTextExtractor
    {
        // Entry point required by the project
        static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_with_formulas.xlsx";

                // Ensure the input file exists; create an empty workbook if it does not.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range whose formulas you want to extract
                const string rangeAddress = "A1:C10";

                // Use Aspose.Cells.Range explicitly to avoid conflict with System.Range
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(rangeAddress);

                // Collect formula texts
                List<string> formulaTexts = new List<string>();

                // Iterate through each cell in the range
                foreach (Cell cell in range)
                {
                    // Retrieve the formula in A1 notation (non‑localized)
                    string formula = cell.GetFormula(false, false);

                    // If the cell contains a formula, store it with its address
                    if (!string.IsNullOrEmpty(formula))
                    {
                        formulaTexts.Add($"{cell.Name}: {formula}");
                    }
                }

                // Output the collected formulas to the console
                Console.WriteLine($"Formulas found in range {rangeAddress}:");
                foreach (string txt in formulaTexts)
                {
                    Console.WriteLine(txt);
                }

                // Write the formulas to a new worksheet for bulk analysis
                int analysisSheetIndex = workbook.Worksheets.Add();
                Worksheet analysisSheet = workbook.Worksheets[analysisSheetIndex];
                analysisSheet.Name = "FormulaAnalysis";

                for (int i = 0; i < formulaTexts.Count; i++)
                {
                    analysisSheet.Cells[i, 0].PutValue(formulaTexts[i]);
                }

                // Save the workbook with the analysis sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
