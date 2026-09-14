// Title: Create a report of cells containing volatile Excel functions after recalculating formulas with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells that recalculates the entire workbook, scans every worksheet for formulas that include volatile functions (e.g., NOW, TODAY, RAND, OFFSET), and writes the worksheet name, cell address, and formula to a new sheet called VolatileReport. | Modify the volatile‑function scanner to accept a custom list of function names supplied at runtime and output the results to a CSV file instead of an Excel worksheet.
// Common Searches: Aspose.Cells C# list cells with volatile formulas after CalculateFormula | how to generate a volatile function report in an Excel file using Aspose.Cells .NET | detect NOW TODAY RAND functions in workbook and export results with Aspose.Cells
// Tags: Aspose.Cells detect volatile functions | C# generate volatile formula report | Aspose.Cells recalculate workbook | Excel volatile function scanning .NET | export volatile cells to worksheet Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, enables the calculation chain, recalculates all formulas, iterates through each worksheet to find cells whose formulas contain any known volatile function, records the sheet name, cell address, and formula, writes this information to a new worksheet named "VolatileReport", and saves the updated workbook as output.xlsx.
class VolatileFunctionsReport
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Enable calculation chain to allow dependency tracking
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Known volatile functions (case‑insensitive)
        string[] volatileFuncs = new string[]
        {
            "NOW()", "TODAY()", "RAND()", "RANDBETWEEN()", "OFFSET()", "INDIRECT()", "INFO()", "CELL()", "AREAS()", "CHOOSE()", "HYPERLINK()", "RTD()", "GETPIVOTDATA()",
            "NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "INFO", "CELL", "AREAS", "CHOOSE", "HYPERLINK", "RTD", "GETPIVOTDATA"
        };

        // Collect cells that contain volatile functions
        List<(string Sheet, string Cell, string Formula)> volatileCells = new List<(string, string, string)>();

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int r = 0; r <= maxRow; r++)
            {
                for (int c = 0; c <= maxCol; c++)
                {
                    Cell cell = cells[r, c];
                    if (cell.IsFormula)
                    {
                        string formulaUpper = cell.Formula?.ToUpperInvariant() ?? "";
                        foreach (string v in volatileFuncs)
                        {
                            if (formulaUpper.Contains(v))
                            {
                                volatileCells.Add((sheet.Name, cell.Name, cell.Formula));
                                break;
                            }
                        }
                    }
                }
            }
        }

        // Create a new worksheet to hold the report
        int reportIndex = workbook.Worksheets.Add();
        Worksheet reportSheet = workbook.Worksheets[reportIndex];
        reportSheet.Name = "VolatileReport";
        Cells reportCells = reportSheet.Cells;

        // Write header
        reportCells["A1"].PutValue("Worksheet");
        reportCells["B1"].PutValue("Cell");
        reportCells["C1"].PutValue("Formula");

        // Populate report rows
        for (int i = 0; i < volatileCells.Count; i++)
        {
            int row = i + 1; // Row 1 is the header
            reportCells[row, 0].PutValue(volatileCells[i].Sheet);
            reportCells[row, 1].PutValue(volatileCells[i].Cell);
            reportCells[row, 2].PutValue(volatileCells[i].Formula);
        }

        // Save the workbook with the report
        workbook.Save("output.xlsx");
    }
}
