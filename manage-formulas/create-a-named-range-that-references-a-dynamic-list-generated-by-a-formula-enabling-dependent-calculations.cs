using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Specify the number of rows for the dynamic SEQUENCE formula
                cells["B1"].PutValue(5); // The dynamic list will have 5 rows

                // 3. Set a dynamic array formula in A1 that spills into neighboring cells
                Cell startCell = cells["A1"];
                startCell.SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

                // 4. Refresh dynamic array formulas so the spill range is materialized
                workbook.RefreshDynamicArrayFormulas(true);

                // 5. Create a named range that refers to the spilled dynamic array (A1#)
                int nameIndex = workbook.Worksheets.Names.Add("DynamicList");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = $"={sheet.Name}!A1#";

                // 6. Use the named range in a dependent calculation (sum of the list)
                cells["C1"].Formula = "=SUM(DynamicList)";

                // 7. Calculate all formulas
                workbook.CalculateFormula();

                // 8. Retrieve the range object via the name to demonstrate it works
                Aspose.Cells.Range namedRange = dynamicName.GetRange();
                Console.WriteLine($"Named range '{dynamicName.Text}' refers to: {namedRange.RefersTo}");
                Console.WriteLine($"Sum of dynamic list (cell C1) = {cells["C1"].Value}");

                // 9. Save the workbook
                string outputPath = "DynamicNamedRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}