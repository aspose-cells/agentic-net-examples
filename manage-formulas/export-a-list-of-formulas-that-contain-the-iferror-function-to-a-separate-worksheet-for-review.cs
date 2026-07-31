// Title: C# Aspose.Cells: Export IFERROR Formulas to a New Worksheet
// Description: Loads a workbook, scans the used range of the first sheet, captures every formula that contains the IFERROR function (case‑insensitive), writes the cell address and formula to a newly added sheet called "IFERROR_Formulas", and saves the updated file.
// Keywords: Aspose.Cells IFERROR export | C# list formulas containing IFERROR | extract error‑handling formulas | create report worksheet Aspose.Cells | scan used range Excel .NET | save workbook with additional sheet | global Excel automation
// Common Searches: how to extract IFERROR formulas with Aspose.Cells | C# code to list cells that use IFERROR | Aspose.Cells create sheet for specific formulas | search workbook for IFERROR function | export formula audit report Aspose.Cells
// Developer Intent: Identify all IFERROR formulas in a workbook and record their addresses and expressions on a separate worksheet.
// Use Cases: Audit spreadsheet error‑handling before release | Generate compliance documentation of IFERROR usage | Quickly locate problematic formulas for troubleshooting | Provide a concise report for finance or data‑validation teams
// AI Prompts: Generate C# Aspose.Cells code that lists cells with the IFERROR function and adds a summary sheet. | Explain how to also capture each formula's evaluated result alongside its address. | Suggest ways to speed up the scan for workbooks with thousands of rows and columns.

using System;
using Aspose.Cells;

namespace AsposeCellsIFERRORExport
{
    // Loads a workbook, scans the used range of the first sheet, captures every formula that contains the IFERROR function (case‑insensitive), writes the cell address and formula to a newly added sheet called "IFERROR_Formulas", and saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (source of formulas)
            Worksheet sourceSheet = workbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Add a new worksheet to hold the IFERROR formulas
            Worksheet reportSheet = workbook.Worksheets.Add("IFERROR_Formulas");
            Cells reportCells = reportSheet.Cells;

            // Write header row in the report sheet
            reportCells[0, 0].PutValue("Cell Address");
            reportCells[0, 1].PutValue("Formula");
            int reportRow = 1; // start after header

            // Determine the used range in the source sheet
            int maxRow = sourceCells.MaxDataRow;
            int maxCol = sourceCells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sourceCells[row, col];

                    // Check if the cell contains a formula with IFERROR (case‑insensitive)
                    if (!string.IsNullOrEmpty(cell.Formula) &&
                        cell.Formula.IndexOf("IFERROR", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Record the cell address and its formula in the report sheet
                        reportCells[reportRow, 0].PutValue(cell.Name);
                        reportCells[reportRow, 1].PutValue(cell.Formula);
                        reportRow++;
                    }
                }
            }

            // Save the workbook with the new report worksheet
            workbook.Save("output.xlsx");
        }
    }
}
