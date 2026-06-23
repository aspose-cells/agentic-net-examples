using System;
using Aspose.Cells;

namespace AsposeCellsArrayNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data that will be used by the dynamic array formula
            // A1:A3 will contain numbers 1, 2, 3
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["A3"].PutValue(3);

            // Set a dynamic array formula in cell B1 that multiplies the range A1:A3 by 2
            // The result will spill into B1:B3
            Cell formulaCell = sheet.Cells["B1"];
            formulaCell.SetDynamicArrayFormula("=A1:A3*2", new FormulaParseOptions(), true);

            // Calculate formulas so that the spilled values are materialized
            workbook.CalculateFormula();

            // Retrieve the spilled range of the dynamic array formula
            // GetArrayRange returns the area that the formula occupies
            CellArea spillArea = formulaCell.GetArrayRange();

            // Build an address string for the spilled range, including the sheet name
            // Example: =Sheet1!$B$1:$B$3
            string startAddress = CellsHelper.CellIndexToName(spillArea.StartRow, spillArea.StartColumn);
            string endAddress = CellsHelper.CellIndexToName(spillArea.EndRow, spillArea.EndColumn);
            string refersTo = $"={sheet.Name}!{startAddress}:{endAddress}";

            // Create a named range that refers to the spilled array result
            int nameIndex = workbook.Worksheets.Names.Add("ArrayResult");
            Name arrayName = workbook.Worksheets.Names[nameIndex];
            arrayName.RefersTo = refersTo; // e.g., =Sheet1!$B$1:$B$3

            // Use the named range in a subsequent calculation, e.g., sum of the array result
            sheet.Cells["C1"].Formula = "=SUM(ArrayResult)";

            // Recalculate to obtain the sum
            workbook.CalculateFormula();

            // Output the sum to the console
            Console.WriteLine("Sum of the dynamic array result (ArrayResult): " + sheet.Cells["C1"].Value);

            // Save the workbook
            workbook.Save("ArrayResultNamedRange.xlsx");
        }
    }
}