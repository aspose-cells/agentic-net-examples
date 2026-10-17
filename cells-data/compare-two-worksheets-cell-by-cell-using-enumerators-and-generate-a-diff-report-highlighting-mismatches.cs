// Title: Generate a cell‑by‑cell diff report between two worksheets in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook with Aspose.Cells, iterates over every cell of the first two worksheets, and records any value mismatches in a newly added worksheet. | Apply a yellow background style to each row in the diff worksheet that contains a mismatched cell. | Modify the comparison logic so that null values and empty strings are treated as equal when building the diff report.
// Common Searches: aspocells compare two worksheets and export differences to a new sheet | c# create excel diff report highlighting mismatched cells using Aspose.Cells | how to iterate over cells in two worksheets with Aspose.Cells and generate a diff worksheet
// Tags: Aspose.Cells worksheet cell comparison | Excel diff report generation C# | highlight mismatched rows Aspose.Cells | enumerate cells Aspose.Cells API | null vs empty string handling Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads a workbook, compares the first two worksheets cell by cell, logs any mismatched cells (address, sheet1 value, sheet2 value) into a new worksheet named "DiffReport", highlights those rows in yellow, and saves the diff report as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string reportPath = "DiffReport.xlsx";

            // Verify source file exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook containing the worksheets to compare.
            Workbook workbook = new Workbook(sourcePath);
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets[1];

            // Add a new worksheet for the diff report.
            int diffSheetIndex = workbook.Worksheets.Add();
            Worksheet diffSheet = workbook.Worksheets[diffSheetIndex];
            diffSheet.Name = "DiffReport";

            // Write header row.
            diffSheet.Cells["A1"].PutValue("Cell Address");
            diffSheet.Cells["B1"].PutValue("Sheet1 Value");
            diffSheet.Cells["C1"].PutValue("Sheet2 Value");

            // Determine the range to compare.
            int maxRow = Math.Max(sheet1.Cells.MaxDataRow, sheet2.Cells.MaxDataRow);
            int maxCol = Math.Max(sheet1.Cells.MaxDataColumn, sheet2.Cells.MaxDataColumn);

            int reportRow = 2; // Data starts after header (1‑based index).

            // Iterate through each cell in the determined range.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    object val1 = sheet1.Cells[row, col].Value;
                    object val2 = sheet2.Cells[row, col].Value;

                    bool equal = (val1 == null && val2 == null) ||
                                 (val1 != null && val1.Equals(val2));

                    if (!equal)
                    {
                        // Record mismatch.
                        string address = CellsHelper.CellIndexToName(row, col);
                        diffSheet.Cells[reportRow, 0].PutValue(address);
                        diffSheet.Cells[reportRow, 1].PutValue(val1);
                        diffSheet.Cells[reportRow, 2].PutValue(val2);

                        // Highlight the row.
                        Style highlightStyle = diffSheet.Cells[reportRow, 0].GetStyle();
                        highlightStyle.ForegroundColor = Color.Yellow;
                        highlightStyle.Pattern = BackgroundType.Solid;
                        for (int c = 0; c < 3; c++)
                        {
                            diffSheet.Cells[reportRow, c].SetStyle(highlightStyle);
                        }

                        reportRow++;
                    }
                }
            }

            // Save the workbook with the diff report.
            workbook.Save(reportPath);
            Console.WriteLine($"Diff report saved to {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
