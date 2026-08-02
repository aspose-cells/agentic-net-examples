using System;
using Aspose.Cells;

namespace AsposeCellsArrayImportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Two‑dimensional data that we want to place into the sheet
            // Each inner array represents a row
            object[][] data = new object[3][];
            data[0] = new object[] { 10, 20, 30 };
            data[1] = new object[] { 40, 50, 60 };
            data[2] = new object[] { 70, 80, 90 };

            // Target cell where the array will start (B2 in this example)
            Cell targetCell = cells["B2"];

            // Set an array formula with pre‑calculated values.
            // The formula itself is not important because we supply the values directly.
            // RowCount = 3, ColumnCount = 3 to match the dimensions of 'data'.
            targetCell.SetArrayFormula(
                "=TRANSPOSE({1,2,3;4,5,6;7,8,9})", // dummy formula; values will be used instead
                rowNumber: 3,
                columnNumber: 3,
                options: new FormulaParseOptions(),
                values: data);

            // Calculate the workbook so that any dependent formulas are refreshed
            workbook.CalculateFormula();

            // Save the result
            workbook.Save("ArrayImportResult.xlsx");
        }
    }
}