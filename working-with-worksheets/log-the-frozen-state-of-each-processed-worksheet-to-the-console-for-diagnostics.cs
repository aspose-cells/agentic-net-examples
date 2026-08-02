// Title: Log Freeze Pane Settings of All Worksheets with Aspose.Cells in C#
// Description: Shows how to create a workbook, set freeze panes on a sheet, and use GetFreezedPanes to read each worksheet’s freeze position, frozen rows and columns, then write the information to the console before saving the file.
// Keywords: Aspose.Cells | GetFreezedPanes | freeze panes | worksheet diagnostics | C# | .NET | console logging | workbook freeze settings
// Common Searches: Aspose.Cells GetFreezedPanes C# example | how to read freeze pane information with Aspose.Cells | log frozen rows and columns for each worksheet .NET | retrieve freeze pane position Aspose.Cells | diagnostic output of freeze panes in C#
// Developer Intent: The developer needs to output the freeze‑pane configuration of every worksheet in a workbook to the console for troubleshooting or verification.
// Use Cases: Confirm that freeze panes are applied correctly before distributing a report. | Create a diagnostic log to investigate scrolling problems in generated Excel files. | Automate unit tests that compare expected frozen rows/columns with actual workbook settings.
// AI Prompts: Write a reusable method that returns a dictionary mapping worksheet names to their freeze pane details using Aspose.Cells. | Generate code that logs freeze pane information to a text file instead of the console for all worksheets in a workbook. | Show how to extend the example to also report split positions when no freeze panes are set.

using System;
using Aspose.Cells;

// Shows how to create a workbook, set freeze panes on a sheet, and use GetFreezedPanes to read each worksheet’s freeze position, frozen rows and columns, then write the information to the console before saving the file.
class LogFrozenState
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets and set freeze panes on one of them
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "FirstSheet";
        ws1.FreezePanes(2, 2, 2, 2); // Freeze at cell C3 with 2 rows and 2 columns frozen

        Worksheet ws2 = workbook.Worksheets.Add("SecondSheet");
        // No freeze panes on ws2

        // Iterate through each worksheet and log its frozen state
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            int row, column, frozenRows, frozenColumns;
            bool hasFreeze = ws.GetFreezedPanes(out row, out column, out frozenRows, out frozenColumns);

            Console.WriteLine($"Worksheet \"{ws.Name}\" (Index {i}):");
            Console.WriteLine($"  Has Freeze Panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"  Freeze Position - Row: {row}, Column: {column}");
                Console.WriteLine($"  Frozen Rows: {frozenRows}, Frozen Columns: {frozenColumns}");
            }
            else
            {
                Console.WriteLine("  No frozen panes.");
            }
        }

        // Save the workbook
        workbook.Save("FrozenStateDemo.xlsx");
    }
}
