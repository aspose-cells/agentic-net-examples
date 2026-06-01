using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data in A1:A5 (unsorted values)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(5 - i); // 5,4,3,2,1
        }

        // Insert a SORT dynamic array formula into cell B2
        Cell targetCell = cells["B2"];
        targetCell.SetDynamicArrayFormula("=SORT(A1:A5)", new FormulaParseOptions(), true);

        // Calculate formulas and refresh dynamic array formulas so the spill range updates
        workbook.CalculateFormula();
        workbook.RefreshDynamicArrayFormulas(true);

        // Retrieve the actual spilled range of the dynamic array formula
        CellArea spillRange = targetCell.GetArrayRange();

        // Output the spilled range coordinates
        Console.WriteLine($"Spill range: Row {spillRange.StartRow + 1} to {spillRange.EndRow + 1}, " +
                          $"Column {spillRange.StartColumn + 1} to {spillRange.EndColumn + 1}");

        // Output the values in the spilled range to verify correct sorting
        for (int row = spillRange.StartRow; row <= spillRange.EndRow; row++)
        {
            Console.WriteLine($"B{row + 1} = {cells[row, spillRange.StartColumn].Value}");
        }

        // Save the workbook
        workbook.Save("SortDynamicArrayDemo.xlsx");
    }
}