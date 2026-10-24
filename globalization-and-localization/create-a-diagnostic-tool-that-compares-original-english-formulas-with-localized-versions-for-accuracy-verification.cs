// Title: Validate Excel formula localization between English and translated workbooks using Aspose.Cells for .NET
// Common Searches: aspnet compare formulas in original and localized Excel files with Aspose.Cells | c# generate report of mismatched formulas between English and translated workbooks | how to verify formula localization accuracy in multi‑language spreadsheets using Aspose.Cells | detect formula differences in localized Excel workbook using C# | create formula mismatch report for translated Excel files with Aspose.Cells .NET
// Tags: Aspose.Cells formula comparison | Excel formula localization validation | C# formula mismatch reporting | compare worksheet formulas Aspose.Cells | detect localized formula differences .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The utility loads an English workbook and its localized counterpart, iterates through matching worksheets, determines a combined used range, and compares cell formulas case‑insensitively after normalizing whitespace. Any mismatched formulas are collected with sheet and cell identifiers, displayed on the console, and written to a new report workbook (FormulaLocalizationReport.xlsx) for further analysis.
class FormulaLocalizationValidator
{
    // Simple container for mismatch information
    class MismatchInfo
    {
        public string SheetName { get; set; } = string.Empty;
        public string CellName { get; set; } = string.Empty;
        public string EnglishFormula { get; set; } = string.Empty;
        public string LocalizedFormula { get; set; } = string.Empty;
    }

    static void Main(string[] args)
    {
        try
        {
            // Paths to the source workbooks
            string englishPath = "OriginalEnglish.xlsx";
            string localizedPath = "LocalizedVersion.xlsx";

            // Verify that input files exist
            if (!File.Exists(englishPath))
                throw new FileNotFoundException($"English workbook not found: {englishPath}");
            if (!File.Exists(localizedPath))
                throw new FileNotFoundException($"Localized workbook not found: {localizedPath}");

            // Load the workbooks
            Workbook englishWb = new Workbook(englishPath);
            Workbook localizedWb = new Workbook(localizedPath);

            // Collect all mismatches
            List<MismatchInfo> mismatches = new List<MismatchInfo>();

            // Compare worksheets one‑by‑one
            int sheetCount = Math.Min(englishWb.Worksheets.Count, localizedWb.Worksheets.Count);
            for (int s = 0; s < sheetCount; s++)
            {
                Worksheet engSheet = englishWb.Worksheets[s];
                Worksheet locSheet = localizedWb.Worksheets[s];

                // Determine the used range for both sheets
                AsposeRange engRange = engSheet.Cells.MaxDisplayRange;
                AsposeRange locRange = locSheet.Cells.MaxDisplayRange;

                // If a sheet is empty, MaxDisplayRange may be null; handle gracefully
                if (engRange == null && locRange == null) continue;
                if (engRange == null) engRange = locRange;
                if (locRange == null) locRange = engRange;

                // Calculate combined boundaries
                int firstRow = Math.Min(engRange.FirstRow, locRange.FirstRow);
                int lastRow = Math.Max(
                    engRange.FirstRow + engRange.RowCount - 1,
                    locRange.FirstRow + locRange.RowCount - 1);
                int firstCol = Math.Min(engRange.FirstColumn, locRange.FirstColumn);
                int lastCol = Math.Max(
                    engRange.FirstColumn + engRange.ColumnCount - 1,
                    locRange.FirstColumn + locRange.ColumnCount - 1);

                // Scan every cell in the combined range
                for (int row = firstRow; row <= lastRow; row++)
                {
                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        Cell engCell = engSheet.Cells[row, col];
                        Cell locCell = locSheet.Cells[row, col];

                        // Compare only when both cells contain formulas
                        if (engCell.IsFormula && locCell.IsFormula)
                        {
                            string? engFormula = engCell.Formula?.Replace(" ", "").Trim();
                            string? locFormula = locCell.Formula?.Replace(" ", "").Trim();

                            // Case‑insensitive comparison of normalized formulas
                            if (!string.Equals(engFormula, locFormula, StringComparison.OrdinalIgnoreCase))
                            {
                                mismatches.Add(new MismatchInfo
                                {
                                    SheetName = engSheet.Name,
                                    CellName = engCell.Name,
                                    EnglishFormula = engCell.Formula ?? string.Empty,
                                    LocalizedFormula = locCell.Formula ?? string.Empty
                                });
                            }
                        }
                    }
                }
            }

            // Output results to console
            if (mismatches.Count == 0)
            {
                Console.WriteLine("All formulas match between the English and localized workbooks.");
            }
            else
            {
                Console.WriteLine("Formula mismatches detected:");
                foreach (var m in mismatches)
                {
                    Console.WriteLine($"Sheet: {m.SheetName}, Cell: {m.CellName}, English: \"{m.EnglishFormula}\", Localized: \"{m.LocalizedFormula}\"");
                }
            }

            // Create a simple report workbook
            Workbook reportWb = new Workbook();
            Worksheet reportSheet = reportWb.Worksheets[0];
            reportSheet.Name = "Formula Mismatches";

            // Header row
            reportSheet.Cells[0, 0].PutValue("Sheet");
            reportSheet.Cells[0, 1].PutValue("Cell");
            reportSheet.Cells[0, 2].PutValue("English Formula");
            reportSheet.Cells[0, 3].PutValue("Localized Formula");

            // Populate report rows
            for (int i = 0; i < mismatches.Count; i++)
            {
                var m = mismatches[i];
                int row = i + 1;
                reportSheet.Cells[row, 0].PutValue(m.SheetName);
                reportSheet.Cells[row, 1].PutValue(m.CellName);
                reportSheet.Cells[row, 2].PutValue(m.EnglishFormula);
                reportSheet.Cells[row, 3].PutValue(m.LocalizedFormula);
            }

            // Save the report
            try
            {
                string reportPath = "FormulaLocalizationReport.xlsx";
                reportWb.Save(reportPath);
                Console.WriteLine($"Report saved to \"{reportPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save report: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
