using System;
using System.Diagnostics;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define size of data set to generate a noticeable calculation load
        int dataRows = 1000;
        int dataCols = 10;

        // Fill cells with numeric values
        for (int row = 0; row < dataRows; row++)
        {
            for (int col = 0; col < dataCols; col++)
            {
                cells[row, col].PutValue(row + col);
            }
        }

        // Add a formula in each row that sums the values of that row
        for (int row = 0; row < dataRows; row++)
        {
            // Example: =SUM(A1:J1) for the first row, =SUM(A2:J2) for the second, etc.
            string sumFormula = $"=SUM(A{row + 1}:{GetColumnLetter(dataCols)}{row + 1})";
            cells[row, dataCols].Formula = sumFormula;
        }

        // -------------------- Automatic mode timing --------------------
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;
        Stopwatch sw = Stopwatch.StartNew();
        workbook.CalculateFormula(); // Calculate all formulas
        sw.Stop();
        Console.WriteLine($"Automatic mode calculation time: {sw.ElapsedMilliseconds} ms");

        // -------------------- Manual mode timing --------------------
        // Switch to manual mode
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Modify a cell to ensure a recalculation is needed
        cells[0, 0].PutValue(999);

        sw.Restart();
        workbook.CalculateFormula(); // Manual calculation invoked explicitly
        sw.Stop();
        Console.WriteLine($"Manual mode calculation time: {sw.ElapsedMilliseconds} ms");

        // Save the workbook (optional, demonstrates usage of the save rule)
        workbook.Save("PerformanceComparison.xlsx");
    }

    // Helper method to convert a 1‑based column index to its Excel column letter (e.g., 1 -> A, 27 -> AA)
    static string GetColumnLetter(int columnNumber)
    {
        int dividend = columnNumber;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        return columnName;
    }
}