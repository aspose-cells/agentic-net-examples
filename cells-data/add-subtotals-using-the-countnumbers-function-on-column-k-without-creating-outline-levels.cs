using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: column J (index 9) as group, column K (index 10) as numbers
        cells["J1"].PutValue("Group");
        cells["K1"].PutValue("Value");

        string[] groups = { "A", "A", "B", "B", "A", "B" };
        int[] values = { 5, 10, 7, 3, 8, 6 };

        for (int i = 0; i < groups.Length; i++)
        {
            cells[i + 1, 9].PutValue(groups[i]);   // Column J
            cells[i + 1, 10].PutValue(values[i]); // Column K
        }

        // Define the range that includes the header and data (J1:K7)
        CellArea area = CellArea.CreateCellArea("J1", "K7");

        // Add subtotals:
        // - Group by the first column of the area (column J) -> offset 0
        // - Use CountNums function on the second column of the area (column K) -> offset 1
        sheet.Cells.Subtotal(area, 0, ConsolidationFunction.CountNums, new int[] { 1 });

        // Save the workbook
        workbook.Save("SubtotalCountNums_K.xlsx");
    }
}