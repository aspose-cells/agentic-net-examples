using System;
using Aspose.Cells;

class DisableSplitPaneDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Demonstrate a split pane before removal
        worksheet.Split();

        // Remove any split window configuration
        worksheet.RemoveSplit();

        // Ensure frozen panes are also cleared, if any
        worksheet.UnFreezePanes();

        // Save the workbook with no split or frozen panes
        workbook.Save("NoSplitPane.xlsx");
    }
}