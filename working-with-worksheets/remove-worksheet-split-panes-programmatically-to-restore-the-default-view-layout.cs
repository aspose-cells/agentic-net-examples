using System;
using Aspose.Cells;

class RemoveSplitDemo
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // OPTIONAL: split the worksheet to simulate an existing split
        sheet.Split();

        // Display pane state before removal (should be Split)
        Console.WriteLine("Pane state before removal: " + sheet.PaneState);

        // Remove any split window
        sheet.RemoveSplit();

        // Ensure frozen panes are also cleared (if any)
        sheet.UnFreezePanes();

        // Display pane state after removal (should be Normal)
        Console.WriteLine("Pane state after removal: " + sheet.PaneState);

        // Save the workbook with the default view layout restored
        workbook.Save("RemovedSplit.xlsx");
    }
}