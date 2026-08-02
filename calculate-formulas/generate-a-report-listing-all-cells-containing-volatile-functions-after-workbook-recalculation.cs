using Aspose.Cells;
using System;
using System.Collections.Generic;

class VolatileFunctionsReport
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample formulas, some of which are volatile
        cells["A1"].Formula = "=NOW()";
        cells["A2"].Formula = "=RAND()";
        cells["A3"].Formula = "=B1+5";
        cells["B1"].Formula = "=INDIRECT(\"A1\")";

        // Enable calculation chain (required for some dependency methods)
        workbook.Settings.FormulaSettings.EnableCalculationChain = true;

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Known volatile functions (lower‑cased for case‑insensitive comparison)
        string[] volatileFunctions = new string[]
        {
            "now()", "today()", "rand()", "randbetween()", "offset()", "indirect()", "cell()", "info()"
        };

        // Collect addresses of cells that contain volatile functions
        List<string> volatileCellAddresses = new List<string>();

        foreach (Cell cell in cells)
        {
            if (!cell.IsFormula) continue; // Skip non‑formula cells

            string formula = cell.Formula?.ToLowerInvariant() ?? string.Empty;

            foreach (string vf in volatileFunctions)
            {
                if (formula.Contains(vf))
                {
                    volatileCellAddresses.Add(cell.Name);
                    break;
                }
            }
        }

        // Create a new worksheet to hold the report
        Worksheet reportSheet = workbook.Worksheets.Add("VolatileReport");
        Cells reportCells = reportSheet.Cells;

        // Header
        reportCells["A1"].PutValue("Cell");
        reportCells["B1"].PutValue("Formula");

        // Populate report rows
        for (int i = 0; i < volatileCellAddresses.Count; i++)
        {
            string address = volatileCellAddresses[i];
            Cell sourceCell = cells[address];

            reportCells[i + 1, 0].PutValue(address);
            reportCells[i + 1, 1].PutValue(sourceCell.Formula);
        }

        // Save the workbook with the report
        workbook.Save("VolatileFunctionsReport.xlsx");
    }
}