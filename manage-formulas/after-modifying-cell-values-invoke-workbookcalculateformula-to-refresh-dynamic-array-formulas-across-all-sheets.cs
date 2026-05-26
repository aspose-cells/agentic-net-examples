using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (in‑memory)
        Workbook wb = new Workbook();

        // Access the first worksheet
        Worksheet ws = wb.Worksheets[0];

        // Add a dynamic array formula that will spill into A1:A5
        // The formula is calculated immediately (third argument = true)
        ws.Cells["A1"].SetDynamicArrayFormula("=SEQUENCE(5)", new FormulaParseOptions(), true);

        // Modify a cell value that could affect other formulas
        ws.Cells["B1"].PutValue(10);

        // After changing cell values, calculate all formulas in the workbook.
        // CalculateFormula also refreshes dynamic array formulas on every sheet.
        wb.CalculateFormula();

        // Display the results of the spilled dynamic array range
        for (int i = 0; i < 5; i++)
        {
            Console.WriteLine($"A{i + 1} = {ws.Cells[i, 0].Value}");
        }

        // Optional: save the workbook to verify the changes
        wb.Save("DynamicArrayRefresh.xlsx");
    }
}