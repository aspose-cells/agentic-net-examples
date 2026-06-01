using System;
using Aspose.Cells;

class HideUnhideColumnDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Add sample data to column B (zero‑based index 1)
        cells["B1"].PutValue("This column will be hidden");
        cells["B2"].PutValue(123);
        cells["B3"].PutValue(456);

        // Hide column B
        cells.HideColumn(1);

        // Save the workbook with the hidden column
        workbook.Save("HiddenColumn.xlsx");

        Console.WriteLine("Column B is hidden. Press 'U' to unhide it, any other key to exit.");
        var key = Console.ReadKey();
        Console.WriteLine();

        if (key.KeyChar == 'U' || key.KeyChar == 'u')
        {
            // Unhide column B; width -1 uses the standard column width
            cells.UnhideColumn(1, -1);

            // Save the workbook after unhiding
            workbook.Save("UnhiddenColumn.xlsx");
            Console.WriteLine("Column B has been unhidden and saved as UnhiddenColumn.xlsx");
        }
        else
        {
            Console.WriteLine("No changes made. The workbook remains with the hidden column.");
        }
    }
}