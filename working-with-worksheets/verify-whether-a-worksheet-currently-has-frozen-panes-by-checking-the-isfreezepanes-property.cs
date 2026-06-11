using Aspose.Cells;
using System;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates a workbook with a default worksheet

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Freeze some panes to demonstrate the check
        // Freeze the first two rows and first two columns
        worksheet.FreezePanes(2, 2, 2, 2);

        // Variables to receive freeze pane details
        int row, column, frozenRows, frozenColumns;

        // GetFreezedPanes returns true if the worksheet has frozen panes
        bool hasFrozenPanes = worksheet.GetFreezedPanes(out row, out column, out frozenRows, out frozenColumns);

        // Output the result
        Console.WriteLine("Worksheet has frozen panes: " + hasFrozenPanes);
        if (hasFrozenPanes)
        {
            Console.WriteLine($"Freeze position - Row: {row}, Column: {column}");
            Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
        }

        // Save the workbook
        workbook.Save("FreezeCheckDemo.xlsx");
    }
}