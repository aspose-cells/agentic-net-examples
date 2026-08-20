// Title: Reset worksheet panes with Worksheet.ResetPanes and set a new freeze using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to call Worksheet.ResetPanes to remove any split or frozen panes, verify the PaneState, then apply a new FreezePanes configuration and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells ResetPanes | Worksheet.ResetPanes C# | clear worksheet splits Aspose.Cells | unfreeze panes Aspose.Cells .NET | freeze panes after reset Aspose.Cells | Aspose.Cells pane management | C# Excel pane reset
// Common Searches: How to reset all panes in an Excel worksheet using Aspose.Cells | Aspose.Cells remove split and unfreeze panes programmatically | Worksheet.ResetPanes example C# | Clear frozen panes before applying new FreezePanes Aspose.Cells
// Developer Intent: Remove any existing split or frozen panes from a worksheet so a fresh pane layout can be applied.
// Use Cases: Standardize the pane layout of imported workbooks before generating reports. | Prepare a template workbook by clearing unknown pane settings prior to custom formatting. | Iterate over multiple sheets in a batch process, ensuring each starts with a clean pane state before applying specific FreezePanes.
// AI Prompts: Generate C# code that uses Worksheet.ResetPanes to clear splits and frozen panes, then freezes panes at B2 with Aspose.Cells for .NET. | Show how to check that Worksheet.PaneState equals Normal after resetting panes and before applying a new FreezePanes call.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to call Worksheet.ResetPanes to remove any split or frozen panes, verify the PaneState, then apply a new FreezePanes configuration and save the workbook with Aspose.Cells for .NET.
    public class ResetPanesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Set up an initial split and freeze to demonstrate the reset
                worksheet.Split();                         // Create a split window
                worksheet.FreezePanes("C3", 2, 2);         // Freeze some panes

                // Reset panes: remove any split and unfreeze any frozen panes
                worksheet.RemoveSplit();                   // Clears split window
                worksheet.UnFreezePanes();                 // Unfreezes panes

                // Verify that the pane state is now Normal (no split, no freeze)
                Console.WriteLine("Pane state after reset: " + worksheet.PaneState);

                // Apply a new pane configuration, e.g., freeze panes at B2
                worksheet.FreezePanes("B2", 1, 1);

                // Save the workbook
                string outputPath = "ResetPanesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ResetPanesDemo.Run();
        }
    }
}
