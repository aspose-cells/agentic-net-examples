// Title: Create an Excel report of cells that use the INDIRECT function with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells list all cells containing INDIRECT formulas in a workbook | how to generate a worksheet report of INDIRECT function usage using Aspose.Cells | extract cell address and formula text for INDIRECT in Excel with Aspose.Cells .NET | scan used range for specific formulas like INDIRECT using Aspose.Cells API | save report of INDIRECT formulas to a new sheet in Aspose.Cells C#
// Tags: Aspose.Cells generate formula usage report | C# extract INDIRECT formulas from Excel | add summary worksheet with formula details Aspose | scan workbook cells for specific function Aspose.Cells | save indirect formula report to new workbook .NET

using System;
using System.IO;
using Aspose.Cells;

namespace IndirectFormulaReport
{
    // The C# program loads an existing workbook, adds a worksheet named "IndirectReport", writes header columns, then iterates through every worksheet (excluding the report sheet) and scans each cell in the used range. When a cell's formula contains the INDIRECT function (case‑insensitive), the cell address and full formula are written to the report sheet. Finally, the workbook is saved to the specified output file, producing a concise report of all INDIRECT formulas in the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Add a new worksheet for the report
                int reportSheetIndex = workbook.Worksheets.Add();
                Worksheet reportSheet = workbook.Worksheets[reportSheetIndex];
                reportSheet.Name = "IndirectReport";

                // Write header row
                reportSheet.Cells["A1"].PutValue("Cell Address");
                reportSheet.Cells["B1"].PutValue("Formula");

                // Row counter for the report (starts after header)
                int reportRow = 1; // zero‑based index; row 1 corresponds to A2/B2

                // Iterate through all worksheets except the report sheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.Equals("IndirectReport", StringComparison.OrdinalIgnoreCase))
                        continue; // skip the report sheet itself

                    Cells cells = sheet.Cells;

                    // Determine the used range
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;

                    // If the sheet is empty, continue
                    if (maxRow < 0 || maxColumn < 0)
                        continue;

                    // Scan each cell in the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Check if the cell contains a formula with the INDIRECT function
                            if (cell.IsFormula && !string.IsNullOrEmpty(cell.Formula) &&
                                cell.Formula.IndexOf("INDIRECT", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                // Write the cell address and its formula to the report sheet
                                reportSheet.Cells[reportRow, 0].PutValue(cell.Name);      // e.g., "B5"
                                reportSheet.Cells[reportRow, 1].PutValue(cell.Formula); // full formula text
                                reportRow++;
                            }
                        }
                    }
                }

                // Save the workbook with the report
                workbook.Save(outputPath);
                Console.WriteLine($"Report generated and saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
