// Title: Extract Excel formulas from a defined range and export them to a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an existing .xlsx file with Aspose.Cells, creates a range (e.g., A1:C10), reads each cell's formula via GetFormula(false,false), collects the formulas in a list, and writes the list to column A of a newly added worksheet before saving. | Show how to iterate over an Aspose.Cells.Range, skip cells without formulas, retrieve the formula text in A1 notation, and export the gathered formulas to a separate sheet in the same workbook.
// Common Searches: aspocells c# extract formulas from a range and save to new worksheet | how to get formula text for multiple cells using Aspose.Cells GetFormula | C# Aspose.Cells bulk formula extraction from A1:C10 | write extracted Excel formulas to another sheet with Aspose.Cells .NET | retrieve non‑empty formulas in a range using Aspose.Cells API
// Tags: Aspose.Cells GetFormula usage | range-based formula retrieval C# | export formula list to new sheet Aspose.Cells | bulk formula analysis .xlsx Aspose.Cells | create and iterate Aspose.Cells range

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormulaExtraction
{
    // The program loads input.xlsx, iterates over the A1:C10 range, captures each cell's formula with GetFormula(false,false), writes the formulas to column A of a newly added worksheet, and saves the workbook as output_with_formulas.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output_with_formulas.xlsx";

                // Verify that the input workbook exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the range from which to extract formulas (e.g., A1:C10)
                string rangeAddress = "A1:C10";

                // Use fully qualified Aspose.Cells.Range to avoid conflict with System.Range
                Aspose.Cells.Range range = cells.CreateRange(rangeAddress);

                // List to hold formula texts for bulk analysis
                List<string> formulaTexts = new List<string>();

                // Iterate through each cell in the defined range
                foreach (Cell cell in range)
                {
                    // The Formula property returns an empty string if there is no formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Get the formula text in A1 notation (isR1C1 = false, isLocal = false)
                        string formula = cell.GetFormula(false, false);
                        formulaTexts.Add(formula);

                        // Optional: output to console for immediate verification
                        Console.WriteLine($"{cell.Name}: {formula}");
                    }
                }

                // Create a new worksheet to store the extracted formulas
                int newSheetIndex = workbook.Worksheets.Add();
                Worksheet analysisSheet = workbook.Worksheets[newSheetIndex];
                Cells analysisCells = analysisSheet.Cells;

                // Write each formula into column A of the new sheet
                for (int i = 0; i < formulaTexts.Count; i++)
                {
                    analysisCells[i, 0].PutValue(formulaTexts[i]);
                }

                // Save the workbook with the analysis sheet added
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
