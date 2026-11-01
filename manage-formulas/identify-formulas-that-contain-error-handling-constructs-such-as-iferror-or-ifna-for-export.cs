// Title: How to locate IFERROR and IFNA formulas in an Excel workbook and export them to a separate sheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find cells with IFERROR formula and export to new worksheet | C# scan Excel workbook for IFNA functions using Aspose.Cells API | extract error handling formulas from a workbook and save to another sheet with Aspose.Cells .NET | list all IFERROR and IFNA formulas in an Excel file using Aspose.Cells | handle missing or password‑protected Excel file while extracting formulas in C#
// Tags: enumerate formulas with Aspose.Cells | extract IFERROR formulas C# | scan workbook cells for IFNA using Aspose.Cells | export error handling formulas to new worksheet | handle password protected workbook Aspose.Cells | save extracted formulas to Excel file C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an input Excel file (gracefully handling missing or password‑protected files), iterates through every worksheet and cell to detect formulas containing IFERROR or IFNA, records each matching cell address and its formula, creates a new worksheet named "ErrorFormulas" with headers, writes the collected data, and saves the workbook to an output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";
        Workbook workbook = null;

        try
        {
            if (File.Exists(inputPath))
            {
                // Attempt to load the workbook; password‑protected files will throw an exception
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Input file not found – create an empty workbook to continue safely
                workbook = new Workbook();
                Console.WriteLine($"Input file '{inputPath}' not found. An empty workbook will be used.");
            }
        }
        catch (Exception ex)
        {
            // Detect password‑protected file by inspecting the exception message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine($"The file '{inputPath}' is password protected and cannot be opened.");
                workbook = new Workbook(); // Continue with an empty workbook
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
        }

        var errorFormulas = new List<string>();

        // Scan all worksheets for formulas containing IFERROR or IFNA
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            var cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cell = cells[row, col];
                    if (cell.IsFormula)
                    {
                        string formula = cell.Formula;
                        if (formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            formula.IndexOf("IFNA", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            errorFormulas.Add($"{sheet.Name}!{cell.Name}: {formula}");
                        }
                    }
                }
            }
        }

        // Add a worksheet to export the collected formulas
        int exportIndex = workbook.Worksheets.Add();
        Worksheet exportSheet = workbook.Worksheets[exportIndex];
        exportSheet.Name = "ErrorFormulas";

        // Header row
        exportSheet.Cells[0, 0].PutValue("Cell");
        exportSheet.Cells[0, 1].PutValue("Formula");

        // Populate export sheet
        for (int i = 0; i < errorFormulas.Count; i++)
        {
            string[] parts = errorFormulas[i].Split(new[] { ": " }, StringSplitOptions.None);
            exportSheet.Cells[i + 1, 0].PutValue(parts[0]); // Cell address
            exportSheet.Cells[i + 1, 1].PutValue(parts.Length > 1 ? parts[1] : string.Empty); // Formula
        }

        // Save the workbook
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
