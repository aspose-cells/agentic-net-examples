using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTest
{
    class AutomaticExceptTableTest
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data in column A (A2:A4)
            cells["A2"].PutValue(10);
            cells["A3"].PutValue(20);
            cells["A4"].PutValue(30);

            // Create a table that includes the source data and a formula column
            // Table range: A1:B4 (A1 header, B1 header, data rows)
            sheet.Cells["A1"].PutValue("Value");
            sheet.Cells["B1"].PutValue("SumFormula");
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Add a formula to the second column of the table (column B)
            // The formula sums the values in column A of the table
            // Using structured reference: =SUM([Value])
            table.PutCellFormula(0, 1, "=SUM([Value])");

            // Set calculation mode to AutomaticExceptTable
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

            // Initial calculation to evaluate the formula
            workbook.CalculateFormula();

            // Capture the initial result of the table formula (first data row)
            Cell formulaCell = sheet.Cells["B2"]; // First row of the table's formula column
            Console.WriteLine("Initial table formula value: " + formulaCell.Value);

            // Change a source value that the table formula depends on
            cells["A2"].PutValue(100); // Change from 10 to 100

            // Without explicit calculation, read the formula cell again
            // In AutomaticExceptTable mode, the table formula should NOT be recalculated automatically
            Console.WriteLine("After source change without recalculation, table formula value: " + formulaCell.Value);

            // Now explicitly recalculate all formulas
            workbook.CalculateFormula();

            // After manual calculation, the table formula should reflect the updated source value
            Console.WriteLine("After manual recalculation, table formula value: " + formulaCell.Value);

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("AutomaticExceptTableTest.xlsx");
        }
    }
}