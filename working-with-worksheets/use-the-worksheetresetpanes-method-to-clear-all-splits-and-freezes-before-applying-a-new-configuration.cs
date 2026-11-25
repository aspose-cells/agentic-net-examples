// Title: Reset worksheet panes and clear splits before applying a new freeze configuration with Aspose.Cells for .NET
// Common Searches: Aspose.Cells reset worksheet panes before setting new freeze panes | C# remove split and frozen panes in Excel workbook using Aspose.Cells | clear all pane splits and freezes Aspose.Cells Worksheet example | reset panes then apply freeze panes Aspose.Cells .NET | how to use RemoveSplit and FreezePanes(0,0,0,0) in Aspose.Cells
// Tags: Worksheet.RemoveSplit method Aspose.Cells | Worksheet.FreezePanes reset parameters Aspose.Cells | reset worksheet panes before new freeze Aspose.Cells | clear splits and frozen panes C# Aspose.Cells | apply new freeze configuration after pane reset Aspose.Cells

using System;
using Aspose.Cells;

namespace ResetPanesDemo
{
    // The example creates a workbook, applies an initial split and freeze, then demonstrates how to clear all splits and frozen panes by calling RemoveSplit and FreezePanes(0,0,0,0). After resetting, a new freeze configuration is applied and the workbook is saved as ResetPanesDemo.xlsx.
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

                // Populate some data so that split/freeze effects are visible
                sheet.Cells["A1"].PutValue("Header");
                for (int i = 1; i <= 30; i++)
                {
                    sheet.Cells[$"A{i + 1}"].PutValue($"Row {i}");
                    sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
                }

                // -----------------------------------------------------------------
                // 2. Apply an initial split and freeze configuration
                // -----------------------------------------------------------------
                Console.WriteLine("Applying initial split and freeze...");

                // Split the window (creates a horizontal/vertical split)
                sheet.Split();

                // Freeze panes at cell C3 (row index 2, column index 2) with 3 rows/columns frozen
                sheet.FreezePanes("C3", 3, 3);

                // Show the pane state after the initial configuration
                Console.WriteLine($"PaneState after initial config: {sheet.PaneState}");

                // -----------------------------------------------------------------
                // 3. Reset all panes (clears splits and freezes)
                // -----------------------------------------------------------------
                Console.WriteLine("Resetting panes to clear all splits and freezes...");

                // Remove any split window
                sheet.RemoveSplit();

                // Clear frozen panes by resetting freeze parameters to zero
                sheet.FreezePanes(0, 0, 0, 0);

                // Verify that the pane state is now Normal (no split/freeze)
                Console.WriteLine($"PaneState after reset: {sheet.PaneState}");

                // -----------------------------------------------------------------
                // 4. Apply a new configuration after reset
                // -----------------------------------------------------------------
                Console.WriteLine("Applying new freeze configuration...");

                // Freeze panes at row 2, column 2 with 2 rows and 2 columns frozen
                sheet.FreezePanes(2, 2, 2, 2);

                // Optionally, split the window again to demonstrate mixed usage
                // sheet.Split(); // Uncomment if you also want a split after freezing

                Console.WriteLine($"PaneState after new config: {sheet.PaneState}");

                // -----------------------------------------------------------------
                // 5. Save the workbook
                // -----------------------------------------------------------------
                string outputPath = "ResetPanesDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Report any unexpected errors – no silent failures
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
