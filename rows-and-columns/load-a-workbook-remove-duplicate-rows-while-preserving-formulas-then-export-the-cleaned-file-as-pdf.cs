// Title: Remove duplicate rows while preserving formulas and export the workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove duplicate rows preserving formulas | How to delete duplicate Excel rows and keep formulas using Aspose.Cells | Export cleaned worksheet to PDF after deduplication with Aspose.Cells .NET | Aspose.Cells C# deduplicate rows based on cell values and formulas | Save workbook as PDF after removing duplicate rows in Aspose.Cells
// Tags: duplicate row removal with formula preservation Aspose.Cells | row signature hashing for deduplication C# | export workbook to PDF Aspose.Cells | delete Excel rows programmatically C# | preserve cell formulas during row deduplication

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, creates a unique signature for each row using its cell formulas and values, deletes rows that share a previously seen signature while adjusting indices, and finally saves the cleaned workbook as a PDF file.
class RemoveDuplicateRowsAndExportPdf
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume processing the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Determine the range of used rows and columns
        int maxRow = cells.MaxDataRow;
        int maxCol = cells.MaxDataColumn;

        // HashSet to store unique row signatures
        HashSet<string> uniqueRows = new HashSet<string>();

        // Iterate through rows
        for (int row = 0; row <= maxRow; row++)
        {
            // Build a signature for the current row based on cell values and formulas
            System.Text.StringBuilder rowSignature = new System.Text.StringBuilder();

            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = cells[row, col];
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Preserve the formula text
                    rowSignature.Append("F:").Append(cell.Formula);
                }
                else if (cell.Value != null)
                {
                    // Preserve the displayed value
                    rowSignature.Append("V:").Append(cell.Value.ToString());
                }
                else
                {
                    rowSignature.Append("E:"); // Empty cell marker
                }
                rowSignature.Append("|"); // Delimiter between cells
            }

            string signature = rowSignature.ToString();

            // Check if this row is a duplicate
            if (uniqueRows.Contains(signature))
            {
                // Delete the duplicate row
                cells.DeleteRow(row);
                // Adjust indices because rows shift up after deletion
                row--;
                maxRow--;
            }
            else
            {
                uniqueRows.Add(signature);
            }
        }

        // Export the cleaned workbook to PDF (replace with your desired output path)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
