using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – set standard height and auto‑fit rows containing formulas
    class StandardHeightAndFormulaAutoFit
    {
        static void Main()
        {
            // Load an existing workbook
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set a new standard row height for the worksheet
            worksheet.Cells.StandardHeight = 20; // example height in points

            // Prepare auto‑fitter options (only affect rows that are not custom‑height)
            AutoFitterOptions options = new AutoFitterOptions
            {
                OnlyAuto = true
            };

            // Iterate through all rows in the worksheet
            foreach (Row row in worksheet.Cells.Rows)
            {
                bool hasFormula = false;

                // Check each cell in the current row for a formula
                foreach (Cell cell in row)
                {
                    if (cell.IsFormula)
                    {
                        hasFormula = true;
                        break;
                    }
                }

                // If the row contains at least one formula, auto‑fit its height
                if (hasFormula)
                {
                    // Auto‑fit a single row using the overload that accepts start row, total rows, and options
                    worksheet.AutoFitRows(row.Index, 1, options);
                }
            }

            // Save the modified workbook
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}