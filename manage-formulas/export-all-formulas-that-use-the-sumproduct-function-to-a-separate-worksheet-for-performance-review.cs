using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportSumProductFormulas
    {
        public static void Main(string[] args)
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
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output_with_sumproduct_review.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook.");
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Add a new worksheet to hold the extracted SUMPRODUCT formulas
            int reviewSheetIndex = workbook.Worksheets.Add();
            Worksheet reviewSheet = workbook.Worksheets[reviewSheetIndex];
            reviewSheet.Name = "SUMPRODUCT Review";

            // Write header row in the review sheet
            Cells reviewCells = reviewSheet.Cells;
            reviewCells["A1"].PutValue("Source Sheet");
            reviewCells["B1"].PutValue("Cell Address");
            reviewCells["C1"].PutValue("Formula");

            int reviewRow = 1; // zero‑based index; start after header

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Skip the review sheet itself to avoid self‑reference
                if (ws.Name == reviewSheet.Name) continue;

                Cells cells = ws.Cells;
                // Enumerate all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula))
                    {
                        // Check if the formula uses the SUMPRODUCT function (case‑insensitive)
                        if (cell.Formula.IndexOf("SUMPRODUCT", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Record the information in the review sheet
                            reviewCells[reviewRow, 0].PutValue(ws.Name);   // Source sheet name
                            reviewCells[reviewRow, 1].PutValue(cell.Name); // Cell address (e.g., "B5")
                            reviewCells[reviewRow, 2].PutValue(cell.Formula); // Full formula text
                            reviewRow++;
                        }
                    }
                }
            }

            // Save the workbook with the new review worksheet
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}