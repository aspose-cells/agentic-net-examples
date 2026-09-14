// Title: C# Aspose.Cells utility to compare English and localized Excel formulas and generate a diagnostic workbook
// AI Prompts: Write C# code that opens two Excel files with Aspose.Cells, iterates every worksheet and cell, and logs any formula differences into a new workbook. | Enhance the comparison routine to identify cells that contain a formula in only one of the workbooks and label them as missing or extra. | Adapt the tool to output the comparison results as a CSV file while preserving worksheet names and cell addresses.
// Common Searches: how to use Aspose.Cells in C# to compare formulas of original and translated Excel workbooks | C# example for generating a formula mismatch report between English and localized spreadsheets | detect missing or extra formulas after localizing Excel files with Aspose.Cells
// Tags: Aspose.Cells compare workbook formulas | Excel formula localization verification C# | generate formula mismatch report Aspose.Cells | detect missing formulas in translated Excel | export comparison results to CSV Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The FormulaLocalizationDiagnostic class loads an original English workbook and a localized workbook using Aspose.Cells, sets the localized workbook’s culture to expose English formulas, iterates through matching worksheets and cells, compares formula presence and text, records mismatches, missing, and extra formulas into a new report workbook, and saves the diagnostic report to the specified path.
public class FormulaLocalizationDiagnostic
{
    /// <param name="originalPath">Path to the original English workbook.</param>
    /// <param name="localizedPath">Path to the localized workbook.</param>
    /// <param name="reportPath">Path where the diagnostic report will be saved.</param>
    public void CompareFormulas(string originalPath, string localizedPath, string reportPath)
    {
        try
        {
            // Verify input files exist.
            if (!File.Exists(originalPath))
                throw new FileNotFoundException($"Original workbook not found: {originalPath}");
            if (!File.Exists(localizedPath))
                throw new FileNotFoundException($"Localized workbook not found: {localizedPath}");

            // Load workbooks.
            Workbook originalWb = new Workbook(originalPath);
            Workbook localizedWb = new Workbook(localizedPath);
            // Ensure localized workbook can expose English formulas.
            localizedWb.Settings.CultureInfo = CultureInfo.CurrentCulture;

            // Create report workbook.
            Workbook reportWb = new Workbook();
            Worksheet reportSheet = reportWb.Worksheets[0];
            Cells reportCells = reportSheet.Cells;

            // Header row.
            reportCells["A1"].PutValue("Worksheet");
            reportCells["B1"].PutValue("Cell Address");
            reportCells["C1"].PutValue("Issue Type");
            reportCells["D1"].PutValue("Original Formula (English)");
            reportCells["E1"].PutValue("Localized Formula (English)");
            int reportRow = 1; // zero‑based index after header.

            // Iterate through matching worksheets.
            int sheetCount = Math.Min(originalWb.Worksheets.Count, localizedWb.Worksheets.Count);
            for (int i = 0; i < sheetCount; i++)
            {
                Worksheet origSheet = originalWb.Worksheets[i];
                Worksheet locSheet = localizedWb.Worksheets[i];
                Cells origCells = origSheet.Cells;
                Cells locCells = locSheet.Cells;

                int maxRow = Math.Max(origCells.MaxDataRow, locCells.MaxDataRow);
                int maxCol = Math.Max(origCells.MaxDataColumn, locCells.MaxDataColumn);

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell origCell = origCells[row, col];
                        Cell locCell = locCells[row, col];
                        bool origHasFormula = origCell.IsFormula;
                        bool locHasFormula = locCell.IsFormula;

                        if (origHasFormula && locHasFormula)
                        {
                            string origFormula = origCell.Formula; // English.
                            string locFormula = locCell.Formula;   // English version of localized formula.

                            if (!string.Equals(origFormula, locFormula, StringComparison.OrdinalIgnoreCase))
                            {
                                reportCells[reportRow, 0].PutValue(origSheet.Name);
                                reportCells[reportRow, 1].PutValue(origCell.Name);
                                reportCells[reportRow, 2].PutValue("Mismatch");
                                reportCells[reportRow, 3].PutValue(origFormula);
                                reportCells[reportRow, 4].PutValue(locFormula);
                                reportRow++;
                            }
                        }
                        else if (origHasFormula && !locHasFormula)
                        {
                            reportCells[reportRow, 0].PutValue(origSheet.Name);
                            reportCells[reportRow, 1].PutValue(origCell.Name);
                            reportCells[reportRow, 2].PutValue("Missing in Localized");
                            reportCells[reportRow, 3].PutValue(origCell.Formula);
                            reportCells[reportRow, 4].PutValue(string.Empty);
                            reportRow++;
                        }
                        else if (!origHasFormula && locHasFormula)
                        {
                            reportCells[reportRow, 0].PutValue(origSheet.Name);
                            reportCells[reportRow, 1].PutValue(locCell.Name);
                            reportCells[reportRow, 2].PutValue("Extra in Localized");
                            reportCells[reportRow, 3].PutValue(string.Empty);
                            reportCells[reportRow, 4].PutValue(locCell.Formula);
                            reportRow++;
                        }
                    }
                }
            }

            // Save the diagnostic report.
            reportWb.Save(reportPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during formula comparison: {ex.Message}");
        }
    }
}

static class Program
{
    static void Main(string[] args)
    {
        // Example usage – adjust file paths as needed.
        var diagnostic = new FormulaLocalizationDiagnostic();
        diagnostic.CompareFormulas("OriginalEnglish.xlsx", "Localized.xlsx", "FormulaComparisonReport.xlsx");
    }
}
