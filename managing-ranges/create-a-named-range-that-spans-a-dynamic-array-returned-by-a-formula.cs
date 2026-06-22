using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Populate some data that the dynamic array formula will reference
        cells["B1"].PutValue(10);
        cells["B2"].PutValue(20);
        cells["B3"].PutValue(30);

        // Set a dynamic array formula in A1 that spills the range B1:B3
        Cell formulaCell = cells["A1"];
        // The method returns the area that the formula should spill into
        CellArea spillArea = formulaCell.SetDynamicArrayFormula("=B1:B3", new FormulaParseOptions(), true);

        // Calculate formulas so the spilled values are materialized
        wb.CalculateFormula();

        // Build the address string of the spilled range (e.g., "B1:B3")
        string startAddr = cells[spillArea.StartRow, spillArea.StartColumn].Name;
        string endAddr   = cells[spillArea.EndRow,   spillArea.EndColumn].Name;
        string rangeAddress = $"{startAddr}:{endAddr}";

        // Create a named range that refers to the spilled dynamic array range
        int nameIdx = wb.Worksheets.Names.Add("MyDynamicArray");
        wb.Worksheets.Names[nameIdx].RefersTo = $"={ws.Name}!{rangeAddress}";

        // Optional: output the RefersTo string to verify
        Console.WriteLine("Named range RefersTo: " + wb.Worksheets.Names[nameIdx].RefersTo);

        // Save the workbook
        wb.Save("DynamicArrayNamedRange.xlsx");
    }
}