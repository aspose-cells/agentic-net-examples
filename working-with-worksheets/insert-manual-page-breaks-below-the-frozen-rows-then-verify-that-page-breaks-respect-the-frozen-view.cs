// Title: Add manual horizontal page breaks below frozen rows and verify they respect the frozen pane using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add horizontal page break after frozen rows | verify page break placement relative to freeze panes using Aspose.Cells | insert manual page breaks below freeze panes in Excel with Aspose.Cells .NET | C# code to set freeze panes then add page breaks in Aspose.Cells workbook | check that horizontal page breaks respect frozen rows in Aspose.Cells
// Tags: horizontal page break insertion Aspose.Cells | freeze panes with Aspose.Cells C# | page break validation after freeze panes | Aspose.Cells workbook pagination control | Excel Xlsx page break handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // The example creates a workbook, populates 20 rows and 5 columns, freezes the top five rows and first two columns, adds horizontal page breaks at rows 6, 12, and 18, checks that each break occurs at or below the first unfrozen row, and saves the file as an XLSX to demonstrate that the page breaks respect the frozen pane.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet.
                // ------------------------------------------------------------
                using (Workbook workbook = new Workbook())
                {
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Name = "Demo";

                    // ------------------------------------------------------------
                    // 2. Populate some sample data (20 rows, 5 columns).
                    // ------------------------------------------------------------
                    const int totalRows = 20;
                    const int totalCols = 5;
                    for (int row = 0; row < totalRows; row++)
                    {
                        for (int col = 0; col < totalCols; col++)
                        {
                            sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                        }
                    }

                    // ------------------------------------------------------------
                    // 3. Freeze the top 5 rows (rows 0‑4) and the first 2 columns (A‑B).
                    //    FreezePanes(row, column, freezedRows, freezedColumns)
                    // ------------------------------------------------------------
                    int freezeRow = 5;   // first unfrozen row (zero‑based)
                    int freezeCol = 2;   // first unfrozen column (zero‑based)
                    sheet.FreezePanes(freezeRow, freezeCol, freezeRow, freezeCol);
                    Console.WriteLine($"Frozen pane set at row {freezeRow} (1‑based: {freezeRow + 1}), column {freezeCol} (1‑based: {freezeCol + 1}).");

                    // ------------------------------------------------------------
                    // 4. Insert manual horizontal page breaks *below* the frozen rows.
                    //    We'll add breaks at rows 6, 12 and 18 (1‑based). The API expects
                    //    zero‑based row indices.
                    // ------------------------------------------------------------
                    int[] breakRowsOneBased = { 6, 12, 18 };
                    foreach (int rowOneBased in breakRowsOneBased)
                    {
                        int zeroBasedRow = rowOneBased - 1;
                        sheet.HorizontalPageBreaks.Add(zeroBasedRow);
                        Console.WriteLine($"Added horizontal page break at row {rowOneBased} (zero‑based index {zeroBasedRow}).");
                    }

                    // ------------------------------------------------------------
                    // 5. Verify that all page breaks are after the frozen rows.
                    // ------------------------------------------------------------
                    bool allBreaksRespectFreeze = true;
                    foreach (HorizontalPageBreak hpb in sheet.HorizontalPageBreaks)
                    {
                        if (hpb.Row < freezeRow)
                        {
                            allBreaksRespectFreeze = false;
                            Console.WriteLine($"Warning: Page break at row {hpb.Row + 1} is above the frozen area (first unfrozen row = {freezeRow + 1}).");
                        }
                        else
                        {
                            Console.WriteLine($"Page break at row {hpb.Row + 1} respects the frozen view.");
                        }
                    }

                    Console.WriteLine(allBreaksRespectFreeze
                        ? "All manual page breaks are positioned below the frozen rows."
                        : "One or more page breaks violate the frozen rows positioning.");

                    // ------------------------------------------------------------
                    // 6. Save the workbook to verify the result in Excel.
                    // ------------------------------------------------------------
                    string outputPath = "PageBreaksDemo.xlsx";

                    // Ensure the directory exists before saving
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{outputPath}'. Open it in Excel to see the frozen pane and page breaks.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
