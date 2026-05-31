using Aspose.Cells;
using Aspose.Cells.Tables;

class RunningTotalRoutine
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data: Item names in column A and amounts in column B
        cells["A1"].PutValue("Item");
        cells["B1"].PutValue("Amount");
        string[] items = { "Item1", "Item2", "Item3", "Item4", "Item5" };
        double[] amounts = { 10, 20, 15, 30, 25 };
        for (int i = 0; i < items.Length; i++)
        {
            cells[i + 1, 0].PutValue(items[i]);   // Column A (zero‑based index 0)
            cells[i + 1, 1].PutValue(amounts[i]); // Column B (zero‑based index 1)
        }

        // Header for the running total column (column C)
        cells["C1"].PutValue("Running Total");

        // Insert running total formula in each row of column C
        // Formula uses absolute reference to the first data cell ($B$2) and expands to the current row
        for (int excelRow = 2; excelRow <= items.Length + 1; excelRow++)
        {
            string formula = $"=SUM($B$2:B{excelRow})";
            // Zero‑based row index is excelRow‑1, column index for C is 2
            cells[excelRow - 1, 2].Formula = formula;
        }

        // Save the workbook
        workbook.Save("RunningTotal.xlsx");
    }
}