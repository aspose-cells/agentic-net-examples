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

        // Set up initial data
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);
        cells["A3"].PutValue(3);

        // Add a formula that depends on the values in column A
        cells["B1"].Formula = "=SUM(A1:A3)";

        // Calculate the initial formulas
        wb.CalculateFormula();

        // Bulk insert 5 rows after the first row (index 1)
        // The third parameter updates references in other formulas automatically
        cells.InsertRows(1, 5, true);

        // Populate the newly inserted rows with new values
        for (int i = 1; i <= 5; i++)
        {
            cells[i, 0].PutValue(i + 10); // Column A values: 11,12,13,14,15
        }

        // Refresh dynamic array formulas if any exist (optional but safe)
        wb.RefreshDynamicArrayFormulas(true);

        // Recalculate all formulas after the bulk insert
        wb.CalculateFormula();

        // Display the updated result of the dependent formula
        Console.WriteLine("Updated B1 value after inserting rows: " + cells["B1"].Value);

        // Save the workbook
        wb.Save("RecalcAfterInsert.xlsx");
    }
}