using System;
using Aspose.Cells;

namespace DynamicArrayNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data that the dynamic array will reference
            // Example: values in B1:B3
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["B3"].PutValue(30);

            // Set a dynamic array formula in A1 that spills the range B1:B3
            // Using SEQUENCE to generate a spill of 3 rows for demonstration
            Cell dynamicCell = cells["A1"];
            CellArea spillArea = dynamicCell.SetDynamicArrayFormula("=B1:B3", new FormulaParseOptions(), true);

            // Calculate formulas and refresh dynamic array spills
            workbook.CalculateFormula();
            workbook.RefreshDynamicArrayFormulas(true);

            // Determine the actual spill range address
            // Use the returned CellArea (may differ if there were obstacles)
            string startAddress = cells[spillArea.StartRow, spillArea.StartColumn].Name;
            string endAddress = cells[spillArea.EndRow, spillArea.EndColumn].Name;
            string spillRef = $"={sheet.Name}!{startAddress}:{endAddress}";

            // Create a named range that refers to the dynamic array spill
            int nameIndex = workbook.Worksheets.Names.Add("SpillRange");
            Name spillName = workbook.Worksheets.Names[nameIndex];
            spillName.RefersTo = spillRef; // e.g., =Sheet1!$A$1:$A$3

            // Use the named range in a subsequent formula (e.g., sum of the spill)
            cells["C1"].Formula = "=SUM(SpillRange)";

            // Recalculate to evaluate the new formula
            workbook.CalculateFormula();

            // Output the result to console
            Console.WriteLine($"Spill range address: {spillRef}");
            Console.WriteLine($"Sum of spill range (C1): {cells["C1"].Value}");

            // Save the workbook (lifecycle save)
            workbook.Save("DynamicArrayNamedRangeDemo.xlsx");
        }
    }
}