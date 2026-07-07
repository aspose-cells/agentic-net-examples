using System;
using Aspose.Cells;

namespace AsposeCellsDynamicArrayNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Input value that will determine the size of the dynamic array
            // Example: B1 = 3 will make SEQUENCE generate 3 rows
            cells["B1"].PutValue(3);

            // Set a dynamic array formula in A1.
            // The formula will spill into neighboring cells based on the value in B1.
            Cell startCell = cells["A1"];
            // SetDynamicArrayFormula returns the range (CellArea) the formula *should* spill into
            CellArea spillArea = startCell.SetDynamicArrayFormula(
                "=SEQUENCE(B1,2)",          // generate a 3‑row, 2‑column array when B1 = 3
                new FormulaParseOptions(), // parsing options (default)
                true);                      // calculate the values immediately

            // Refresh dynamic array formulas so that the spill range is materialized
            wb.RefreshDynamicArrayFormulas(true);

            // Build the address string for the spilled range (e.g., Sheet1!A1:B3)
            string startAddress = cells[spillArea.StartRow, spillArea.StartColumn].Name;
            string endAddress   = cells[spillArea.EndRow,   spillArea.EndColumn].Name;
            string rangeAddress = $"{sheet.Name}!{startAddress}:{endAddress}";

            // Create a named range that refers to the spilled dynamic array
            int nameIdx = wb.Worksheets.Names.Add("MyDynamicArray");
            Name dynName = wb.Worksheets.Names[nameIdx];
            dynName.RefersTo = $"={rangeAddress}";

            // Optional: verify the named range points to the correct area
            Console.WriteLine($"Named range '{dynName.Text}' refers to: {dynName.RefersTo}");

            // Save the workbook
            wb.Save("DynamicArrayNamedRange.xlsx");
        }
    }
}