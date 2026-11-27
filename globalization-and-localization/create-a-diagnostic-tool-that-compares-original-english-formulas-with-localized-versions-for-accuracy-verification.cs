// Title: C# Aspose.Cells utility to validate English versus localized Excel formulas and create a mismatch report
// Common Searches: Aspose.Cells C# compare formulas in original and localized Excel workbooks | how to generate a formula translation mismatch report with Aspose.Cells | C# code to verify Excel formula localization accuracy using Aspose.Cells | detect formula differences between English and localized Excel files in .NET
// Tags: formula equivalence check Aspose.Cells | localization mismatch report generation | worksheet formula comparison .NET | Excel formula translation validation C# | Aspose.Cells workbook diagnostic tool

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace FormulaLocalizationDiagnostic
{
    // The sample loads an English source workbook and its localized counterpart with Aspose.Cells, walks through matching worksheets and used cell ranges, compares each pair of formulas using a whitespace‑ and case‑insensitive equivalence test, records any mismatches, and writes an Excel report that lists the sheet name, cell address, English formula, and localized formula.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the workbooks (replace with actual file locations)
            string originalPath = @"C:\Docs\OriginalEnglish.xlsx";
            string localizedPath = @"C:\Docs\LocalizedVersion.xlsx";
            string reportPath = @"C:\Docs\LocalizationReport.xlsx";

            try
            {
                // Verify input files exist
                if (!File.Exists(originalPath))
                {
                    Console.WriteLine($"Original workbook not found: {originalPath}");
                    return;
                }

                if (!File.Exists(localizedPath))
                {
                    Console.WriteLine($"Localized workbook not found: {localizedPath}");
                    return;
                }

                // Load the original (English) workbook
                Workbook originalWb = new Workbook(originalPath);
                // Load the localized workbook
                Workbook localizedWb = new Workbook(localizedPath);

                // Prepare a list to hold mismatch information
                List<string> mismatches = new List<string>();

                // Iterate through each worksheet (assumes same worksheet order and count)
                int sheetCount = Math.Min(originalWb.Worksheets.Count, localizedWb.Worksheets.Count);
                for (int s = 0; s < sheetCount; s++)
                {
                    Worksheet origSheet = originalWb.Worksheets[s];
                    Worksheet locSheet = localizedWb.Worksheets[s];

                    // Determine the used range to limit iteration
                    var usedRange = origSheet.Cells.MaxDisplayRange; // Returns a Range object

                    int startRow = usedRange.FirstRow;
                    int endRow = startRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = startCol + usedRange.ColumnCount - 1;

                    // Iterate through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell origCell = origSheet.Cells[row, col];
                            Cell locCell = locSheet.Cells[row, col];

                            // Only process cells that contain formulas
                            if (!string.IsNullOrEmpty(origCell.Formula) && !string.IsNullOrEmpty(locCell.Formula))
                            {
                                string englishFormula = origCell.Formula;   // English (internal) formula
                                string localizedFormula = locCell.Formula; // Localized formula as stored

                                // Compare the formulas; if they differ, record the mismatch
                                if (!AreFormulasEquivalent(englishFormula, localizedFormula))
                                {
                                    string address = origCell.Name;
                                    mismatches.Add($"Sheet: {origSheet.Name}, Cell: {address}, English: \"{englishFormula}\", Localized: \"{localizedFormula}\"");
                                }
                            }
                        }
                    }
                }

                // Generate a simple report workbook
                Workbook reportWb = new Workbook();
                Worksheet reportSheet = reportWb.Worksheets[0];
                reportSheet.Name = "Localization Mismatches";

                // Header
                reportSheet.Cells[0, 0].PutValue("Sheet");
                reportSheet.Cells[0, 1].PutValue("Cell");
                reportSheet.Cells[0, 2].PutValue("English Formula");
                reportSheet.Cells[0, 3].PutValue("Localized Formula");

                // Populate report rows
                for (int i = 0; i < mismatches.Count; i++)
                {
                    // Split the stored string to extract components
                    // Format: Sheet: {sheet}, Cell: {cell}, English: "{eng}", Localized: "{loc}"
                    string entry = mismatches[i];
                    string[] parts = entry.Split(new[] { ", " }, StringSplitOptions.None);

                    string sheetPart = parts[0].Substring("Sheet: ".Length);
                    string cellPart = parts[1].Substring("Cell: ".Length);
                    string engPart = parts[2].Substring("English: \"".Length);
                    engPart = engPart.Substring(0, engPart.Length - 1); // remove trailing quote
                    string locPart = parts[3].Substring("Localized: \"".Length);
                    locPart = locPart.Substring(0, locPart.Length - 1); // remove trailing quote

                    int rowIdx = i + 1;
                    reportSheet.Cells[rowIdx, 0].PutValue(sheetPart);
                    reportSheet.Cells[rowIdx, 1].PutValue(cellPart);
                    reportSheet.Cells[rowIdx, 2].PutValue(engPart);
                    reportSheet.Cells[rowIdx, 3].PutValue(locPart);
                }

                // Auto-fit columns for readability
                reportSheet.AutoFitColumns();

                // Ensure the directory for the report exists
                string reportDir = Path.GetDirectoryName(reportPath);
                if (!string.IsNullOrEmpty(reportDir) && !Directory.Exists(reportDir))
                {
                    Directory.CreateDirectory(reportDir);
                }

                // Save the report
                reportWb.Save(reportPath, SaveFormat.Xlsx);

                Console.WriteLine($"Localization diagnostic completed. {mismatches.Count} mismatches found.");
                Console.WriteLine($"Report saved to: {reportPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
            }
        }

        // Simple equivalence check – can be extended for deeper analysis
        private static bool AreFormulasEquivalent(string englishFormula, string localizedFormula)
        {
            // For a quick check, compare after normalizing whitespace and case.
            string normEng = englishFormula.Replace(" ", "").ToUpperInvariant();
            string normLoc = localizedFormula.Replace(" ", "").ToUpperInvariant();

            return normEng == normLoc;
        }
    }
}
