// Title: C# – Log Frozen Pane Settings of All Worksheets with Aspose.Cells
// Description: Shows how to create a workbook, apply frozen panes to selected sheets, and use GetFreezedPanes to detect each worksheet’s freeze position, frozen rows and columns, then write the information to the console before saving the file.
// Keywords: Aspose.Cells | C# | .NET | GetFreezedPanes | frozen panes | worksheet freeze state | FreezePanes method | diagnostic logging | console output | retrieve frozen rows columns
// Common Searches: Aspose.Cells get frozen pane status C# | How to check if a worksheet has frozen panes using Aspose.Cells | Log freeze pane coordinates for each sheet in a workbook | Retrieve frozen rows and columns count with Aspose.Cells .NET | Console diagnostics for frozen panes Aspose.Cells
// Developer Intent: Extract and display the freeze‑pane configuration of every worksheet in a workbook.
// Use Cases: Verify that required panes are frozen before publishing a spreadsheet. | Create a diagnostic report of freeze settings to troubleshoot layout issues. | Conditionally apply formatting or calculations based on the presence of frozen rows or columns.
// AI Prompts: Generate C# code that iterates through all worksheets in an Aspose.Cells workbook and writes each sheet's frozen pane details to a text file. | Explain how the GetFreezedPanes method works and describe the meaning of its out parameters (row, column, frozenRows, frozenColumns). | Show how to change the frozen pane position of a specific worksheet after logging its current settings.

using Aspose.Cells;
using System;

// Shows how to create a workbook, apply frozen panes to selected sheets, and use GetFreezedPanes to detect each worksheet’s freeze position, frozen rows and columns, then write the information to the console before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // First worksheet with frozen panes at C3 (row 2, column 2) and 2 frozen rows/columns
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.FreezePanes(2, 2, 2, 2);

        // Second worksheet without frozen panes
        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");

        // Third worksheet with frozen panes at D5 (row 4, column 3) and 4 frozen rows/columns
        Worksheet ws3 = workbook.Worksheets.Add("Sheet3");
        ws3.FreezePanes("D5", 4, 4);

        // Log frozen state for each worksheet
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            int row, column, frozenRows, frozenColumns;
            bool hasFreeze = ws.GetFreezedPanes(out row, out column, out frozenRows, out frozenColumns);
            Console.WriteLine($"Worksheet '{ws.Name}': Frozen = {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"  Freeze Position - Row: {row}, Column: {column}");
                Console.WriteLine($"  Frozen Rows: {frozenRows}, Frozen Columns: {frozenColumns}");
            }
        }

        // Save the workbook
        workbook.Save("FrozenStateDemo.xlsx");
    }
}
