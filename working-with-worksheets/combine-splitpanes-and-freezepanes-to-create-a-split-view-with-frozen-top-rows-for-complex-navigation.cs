// Title: Create a split view with frozen top rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# split pane and freeze top rows example | C# code to split Excel window and freeze first two rows with Aspose.Cells | How to use Split() and FreezePanes together in Aspose.Cells .NET | Create split view with frozen header rows in Excel via Aspose.Cells C# | Set active cell for split and then freeze panes in an Aspose.Cells workbook
// Tags: Aspose.Cells split pane API C# | Aspose.Cells freeze panes after split | Excel split view with frozen header rows Aspose | C# workbook split and freeze example | SetActiveCell for split view Aspose.Cells

using System;
using Aspose.Cells;

namespace SplitAndFreezeDemo
{
    // The sample creates a new workbook, fills 100 rows and 5 columns with data, sets the active cell to E5, calls Split() to divide the window, then freezes the first two rows in the top‑left pane using FreezePanes("A3", 2, 0), and finally saves the file as SplitAndFreezeDemo.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DemoSheet";

                // -----------------------------------------------------------------
                // 2. Populate the worksheet with sample data (100 rows, 5 columns)
                // -----------------------------------------------------------------
                for (int row = 0; row < 100; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        // Header row
                        if (row == 0)
                        {
                            sheet.Cells[row, col].PutValue($"Header {col + 1}");
                        }
                        else
                        {
                            sheet.Cells[row, col].PutValue($"R{row}C{col + 1}");
                        }
                    }
                }

                // -----------------------------------------------------------------
                // 3. Set the active cell where the split will occur.
                //    The Split() method splits the window at the active cell.
                // -----------------------------------------------------------------
                sheet.ActiveCell = "E5"; // Split after column D and row 4
                sheet.Split();           // Perform the split

                // -----------------------------------------------------------------
                // 4. Freeze the top rows (e.g., first 2 rows) in the top‑left pane.
                //    FreezePanes works on the active pane after a split.
                // -----------------------------------------------------------------
                // Freeze rows above cell A3 (i.e., rows 1‑2). No columns are frozen.
                sheet.FreezePanes("A3", 2, 0);

                // -----------------------------------------------------------------
                // 5. Optional: Adjust the view so the split line is visible.
                // -----------------------------------------------------------------
                // Set the first visible row/column for each pane if desired.
                sheet.FirstVisibleRow = 0;
                sheet.FirstVisibleColumn = 0;

                // -----------------------------------------------------------------
                // 6. Save the workbook to a file
                // -----------------------------------------------------------------
                string outputPath = "SplitAndFreezeDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                Console.WriteLine("The worksheet now has a split view with the top two rows frozen.");
            }
            catch (Exception ex)
            {
                // Any exception will be printed to the console – no failures are hidden.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
