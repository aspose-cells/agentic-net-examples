// Title: C# Diagnostic Tool to Compare English Formula and FormulaLocal in Excel with Aspose.Cells
// Description: A C# program that loads an Excel workbook, adds a "FormulaComparison" sheet, scans every cell, extracts the standard English formula (Formula) and the locale‑specific formula (FormulaLocal), checks for equality, records the cell address, both formulas, and a match flag, then saves the workbook with the diagnostic report.
// Keywords: Aspose.Cells | C# | .NET | Formula | FormulaLocal | Excel formula localization | globalization | localization QA | diagnostic report | workbook comparison
// Common Searches: Aspose.Cells compare Formula and FormulaLocal | C# generate formula localization report | verify localized formulas in Excel using Aspose.Cells | Excel formula localization diagnostic .NET | check FormulaLocal vs Formula with Aspose.Cells
// Developer Intent: Generate a workbook that lists each formula cell and shows whether its English and localized formulas match.
// Use Cases: Detect translation errors where FormulaLocal differs from the original Formula during localization QA. | Provide language engineers with a ready‑to‑review report of formula consistency across all worksheets. | Automate pre‑release verification of formula localization in Excel templates.
// AI Prompts: Show how to highlight mismatched cells with a red background in the diagnostic report. | Provide code to export the comparison results to a CSV file instead of adding a new worksheet. | Explain how to handle cells that raise errors when comparing Formula and FormulaLocal while still logging the discrepancy.

using System;
using System.IO;
using Aspose.Cells;

namespace FormulaLocalizationDiagnostic
{
    // A C# program that loads an Excel workbook, adds a "FormulaComparison" sheet, scans every cell, extracts the standard English formula (Formula) and the locale‑specific formula (FormulaLocal), checks for equality, records the cell address, both formulas, and a match flag, then saves the workbook with the diagnostic report.
    public class DiagnosticTool
    {
        private readonly string _inputPath;
        private readonly string _outputPath;

        public DiagnosticTool(string inputPath, string outputPath)
        {
            _inputPath = inputPath;
            _outputPath = outputPath;
        }

        public void Run()
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(_inputPath))
                    throw new FileNotFoundException($"Input workbook not found: {_inputPath}");

                // Ensure output directory exists
                string outDir = Path.GetDirectoryName(_outputPath);
                if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    Directory.CreateDirectory(outDir);

                // Load workbook (lifecycle rule: load)
                Workbook workbook = new Workbook(_inputPath);

                // Add report worksheet
                int reportIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportIndex];
                reportSheet.Name = "FormulaComparison";

                // Write header row
                reportSheet.Cells["A1"].PutValue("Cell");
                reportSheet.Cells["B1"].PutValue("English Formula (Formula)");
                reportSheet.Cells["C1"].PutValue("Localized Formula (FormulaLocal)");
                reportSheet.Cells["D1"].PutValue("Match?");

                int reportRow = 1; // zero‑based index (row 2 in Excel)

                // Iterate through all worksheets and cells
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the report sheet itself
                    if (ws.Name == "FormulaComparison") continue;

                    int maxRow = ws.Cells.MaxDataRow;
                    int maxCol = ws.Cells.MaxDataColumn;

                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            Cell cell = ws.Cells[row, col];
                            if (cell.IsFormula)
                            {
                                string englishFormula = cell.Formula;          // Standard (en‑US) formula
                                string localizedFormula = cell.FormulaLocal;   // Locale‑specific formula

                                // Simple verification: if both strings are equal after trimming, consider it a match.
                                bool isMatch = string.Equals(
                                    englishFormula?.Trim(),
                                    localizedFormula?.Trim(),
                                    StringComparison.OrdinalIgnoreCase);

                                // Write the comparison result to the report sheet
                                reportSheet.Cells[reportRow, 0].PutValue($"{ws.Name}!{cell.Name}");
                                reportSheet.Cells[reportRow, 1].PutValue(englishFormula);
                                reportSheet.Cells[reportRow, 2].PutValue(localizedFormula);
                                reportSheet.Cells[reportRow, 3].PutValue(isMatch ? "Yes" : "No");

                                reportRow++;
                            }
                        }
                    }
                }

                // Save the workbook with the diagnostic report (lifecycle rule: save)
                workbook.Save(_outputPath);
                Console.WriteLine($"Diagnostic report saved to {_outputPath}");
            }
            catch (Exception ex)
            {
                // Handle exceptions (log, rethrow, or display as needed)
                Console.Error.WriteLine($"Error during diagnostic run: {ex.Message}");
                throw;
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                string inputPath;
                string outputPath;

                if (args.Length >= 2)
                {
                    inputPath = args[0];
                    outputPath = args[1];
                }
                else
                {
                    // Default paths for quick testing
                    inputPath = "input.xlsx";
                    outputPath = "output_with_report.xlsx";
                }

                var tool = new DiagnosticTool(inputPath, outputPath);
                tool.Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
