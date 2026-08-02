using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set calculation mode to AutomaticExceptTable.
            // Table formulas will be recalculated by Excel,
            // while other formulas stay static unless CalculateFormula is called.
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // -------------------------------------------------
            // Add a regular (non‑table) formula – it will NOT recalc automatically.
            // -------------------------------------------------
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue(5);
            ws.Cells["B1"].Formula = "=A1*2"; // static unless workbook.CalculateFormula() is invoked

            // -------------------------------------------------
            // Create a simple table with a calculated column.
            // Table formulas are subject to the AutomaticExceptTable mode.
            // -------------------------------------------------
            int startRow = 2; // Excel rows are 1‑based; start at row 2 for the table data

            ws.Cells[$"A{startRow}"].PutValue(1);
            ws.Cells[$"A{startRow + 1}"].PutValue(2);
            ws.Cells[$"A{startRow + 2}"].PutValue(3);

            // Add a ListObject (table) covering A{startRow}:A{startRow+2}
            // Parameters: firstRow (0‑based), firstColumn, totalRows, totalColumns, hasHeaders
            int firstRowIndex = startRow - 1; // convert to zero‑based index
            int firstColumnIndex = 0;
            int totalRows = 3;
            int totalColumns = 1;
            bool hasHeaders = true;

            // Create the table
            ws.ListObjects.Add(firstRowIndex, firstColumnIndex,
                               firstRowIndex + totalRows - 1,
                               firstColumnIndex + totalColumns - 1,
                               hasHeaders);

            // Add a calculated column inside the table (column B)
            // This formula will be recalculated by Excel when the workbook is opened.
            ws.Cells[$"B{startRow}"].Formula = $"=A{startRow}*10";

            // Save the workbook
            string outputPath = "AutomaticExceptTable.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}