using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the worksheet tab color (e.g., LightGreen)
        sheet.TabColor = Color.LightGreen;

        // Freeze the first column (A) while keeping all rows scrollable
        // Freeze at column index 1 (B) with 0 rows frozen and 1 column frozen
        sheet.FreezePanes(0, 1, 0, 1);

        // Save the workbook
        workbook.Save("TabColorAndFreezeFirstColumn.xlsx");
    }
}