// Title: Check if a worksheet has frozen panes using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to detect frozen panes in a worksheet | C# example using GetFreezedPanes to read frozen rows and columns | Determine if worksheet panes are frozen or split with Aspose.Cells PaneState | Verify freeze panes status in an Excel file using Aspose.Cells .NET | Read frozen pane coordinates from a workbook with Aspose.Cells API
// Tags: GetFreezedPanes Aspose.Cells C# | PaneState property Aspose.Cells | detect frozen panes Aspose.Cells | worksheet freeze panes verification .NET | read frozen rows columns Aspose.Cells

using System;
using Aspose.Cells;

namespace VerifyFreezePanesDemo
{
    // The sample loads or creates a workbook, accesses the first worksheet, and determines whether panes are frozen by calling GetFreezedPanes and inspecting the PaneState property, then outputs the freeze status and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Determine the workbook to work with.
            // If a file path is supplied as the first argument, load that workbook.
            // Otherwise, create a new workbook and optionally freeze panes for demonstration.
            Workbook workbook = null;
            string workbookPath = null;

            try
            {
                if (args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]))
                {
                    workbookPath = args[0];
                    Console.WriteLine($"Loading workbook from \"{workbookPath}\"...");
                    workbook = new Workbook(workbookPath);
                }
                else
                {
                    Console.WriteLine("No input file supplied. Creating a new workbook for demonstration.");
                    workbook = new Workbook();
                    // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns.
                    Worksheet demoSheet = workbook.Worksheets[0];
                    demoSheet.Name = "DemoSheet";
                    demoSheet.FreezePanes(2, 2, 2, 2);
                    // Save the demo workbook so the user can inspect it later.
                    workbookPath = "DemoWorkbook.xlsx";
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Demo workbook saved as \"{workbookPath}\" with frozen panes.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading or creating workbook: {ex.Message}");
                return;
            }

            // At this point we have a valid workbook instance.
            Worksheet worksheet = workbook.Worksheets[0];
            Console.WriteLine($"Checking worksheet \"{worksheet.Name}\" for frozen panes...");

            // Method 1: Use GetFreezedPanes which returns a bool indicating whether panes are frozen.
            bool hasFreeze = false;
            int freezeRow = 0, freezeColumn = 0, frozenRows = 0, frozenColumns = 0;

            try
            {
                hasFreeze = worksheet.GetFreezedPanes(out freezeRow, out freezeColumn, out frozenRows, out frozenColumns);
                Console.WriteLine($"GetFreezedPanes returned: {hasFreeze}");
                if (hasFreeze)
                {
                    Console.WriteLine($"Freeze position - Row: {freezeRow}, Column: {freezeColumn}");
                    Console.WriteLine($"Number of frozen rows: {frozenRows}, Number of frozen columns: {frozenColumns}");
                }
                else
                {
                    Console.WriteLine("No frozen panes detected via GetFreezedPanes.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing GetFreezedPanes: {ex.Message}");
            }

            // Method 2: Use the PaneState property. If the state is not PaneStateType.None, panes are either split or frozen.
            try
            {
                PaneStateType paneState = worksheet.PaneState;
                Console.WriteLine($"Worksheet.PaneState = {paneState}");
                if (paneState == PaneStateType.Frozen)
                {
                    Console.WriteLine("PaneState indicates that panes are frozen.");
                }
                else if (paneState == PaneStateType.Split)
                {
                    Console.WriteLine("PaneState indicates that panes are split (not frozen).");
                }
                else
                {
                    Console.WriteLine("PaneState indicates no split or freeze.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error accessing PaneState: {ex.Message}");
            }

            // Final status summary
            Console.WriteLine();
            Console.WriteLine("=== Summary ===");
            Console.WriteLine($"Worksheet \"{worksheet.Name}\" frozen panes: {hasFreeze}");
            Console.WriteLine($"PaneState: {worksheet.PaneState}");

            // Optionally, save any changes (e.g., if we created a demo workbook).
            try
            {
                if (!string.IsNullOrEmpty(workbookPath))
                {
                    workbook.Save(workbookPath);
                    Console.WriteLine($"Workbook saved to \"{workbookPath}\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
