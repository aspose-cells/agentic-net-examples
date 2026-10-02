// Title: C# – Remove duplicate rows (keeping formulas) from an Excel worksheet and save as PDF using Aspose.Cells
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, builds a signature for each row using displayed values and formula text, deletes rows that duplicate an existing signature while preserving formulas, and then saves the workbook as a PDF. | Write a method in .NET that uses Aspose.Cells to identify and remove duplicate rows from a worksheet, ensures any formulas remain intact, and exports the cleaned sheet to a PDF file.
// Common Searches: Aspose.Cells C# how to delete duplicate rows while keeping formulas | remove duplicate rows from Excel workbook using Aspose.Cells and export to PDF | preserve Excel formulas when deduplicating rows with Aspose.Cells .NET | save cleaned worksheet as PDF after removing duplicate rows Aspose.Cells
// Tags: duplicate row detection Aspose.Cells | formula preservation Aspose.Cells | PDF export Aspose.Cells | bottom‑up row deletion Aspose.Cells | row signature hashing Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads 'input.xlsx' with Aspose.Cells, creates a unique signature for each row based on cell values and formulas, removes rows that share the same signature (processing deletions from bottom to top), and saves the resulting workbook as 'output.pdf' in PDF format.
class Program
{
    static void Main()
    {
        // Load the workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the used range
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // HashSet to track unique row signatures
        HashSet<string> seenRows = new HashSet<string>();
        // List to collect indices of duplicate rows
        List<int> duplicateRowIndices = new List<int>();

        // Build a signature for each row (including formulas) and detect duplicates
        for (int row = 0; row <= maxRow; row++)
        {
            List<string> parts = new List<string>();
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (cell == null)
                {
                    parts.Add(string.Empty);
                }
                else if (cell.IsFormula)
                {
                    // Preserve formula text in the signature
                    parts.Add(cell.Formula);
                }
                else
                {
                    // Use displayed value for non‑formula cells
                    parts.Add(cell.Value?.ToString() ?? string.Empty);
                }
            }

            string rowKey = string.Join("|", parts);
            if (!seenRows.Add(rowKey))
            {
                // Row already exists → mark for deletion
                duplicateRowIndices.Add(row);
            }
        }

        // Delete duplicate rows from bottom to top to keep indices valid
        for (int i = duplicateRowIndices.Count - 1; i >= 0; i--)
        {
            sheet.Cells.DeleteRow(duplicateRowIndices[i]);
        }

        // Export the cleaned workbook as PDF (save rule)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
