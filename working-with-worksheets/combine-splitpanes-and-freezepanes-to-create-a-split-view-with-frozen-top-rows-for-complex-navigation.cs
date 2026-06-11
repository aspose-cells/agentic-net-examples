using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate the worksheet with sample data
        for (int i = 0; i < 100; i++)
        {
            for (int j = 0; j < 20; j++)
            {
                sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
            }
        }

        // Freeze the top 4 rows (rows 0‑3) while keeping all columns unfrozen
        // Parameters: row index, column index, number of frozen rows, number of frozen columns
        sheet.FreezePanes(4, 0, 4, 0);

        // Split the window to create separate panes
        sheet.Split();

        // Adjust the bottom pane so that it starts just after the frozen rows
        PaneCollection panes = sheet.GetPanes();
        panes.FirstVisibleRowOfBottomPane = 4; // first visible row in the bottom pane

        // Save the workbook
        workbook.Save("SplitFreezeDemo.xlsx");
    }
}