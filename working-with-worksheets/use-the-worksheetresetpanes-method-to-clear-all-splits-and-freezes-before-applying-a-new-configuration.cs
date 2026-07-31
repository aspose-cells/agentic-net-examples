// Title: Clear splits and frozen panes with Worksheet.ResetPanes before applying FreezePanes in Aspose.Cells for .NET
// Description: C# example that defines a WorksheetExtensions.ResetPanes method (RemoveSplit + UnFreezePanes) to clear all pane settings, then freezes panes at cell C3 and saves the workbook as ResetPanesDemo.xlsx.
// Keywords: Aspose.Cells ResetPanes | Worksheet.RemoveSplit | Worksheet.UnFreezePanes | C# clear frozen panes | FreezePanes example | pane management Aspose.Cells | .NET worksheet panes
// Common Searches: Aspose.Cells reset panes before freeze | C# Worksheet.ResetPanes usage | remove split window Aspose.Cells | unfreeze panes programmatically .NET | how to clear frozen panes in Aspose.Cells
// Developer Intent: Remove any existing split or frozen panes from a worksheet prior to setting a new FreezePanes configuration.
// Use Cases: Standardize workbook layout by clearing previous pane settings before applying a specific freeze. | Refresh a template workbook where old splits must be removed before generating a report. | Update user‑driven view preferences by resetting panes and then applying a new freeze position.
// AI Prompts: Write C# code using Aspose.Cells to clear all splits and frozen panes, then freeze panes at cell D5. | Explain the internal actions performed by Worksheet.ResetPanes in Aspose.Cells. | Create a step‑by‑step guide for building an extension method that resets panes and applies a custom FreezePanes range.

using System;
using System.IO;
using Aspose.Cells;

namespace ResetPanesExample
{
    // Extension method to clear splits and frozen panes
    // C# example that defines a WorksheetExtensions.ResetPanes method (RemoveSplit + UnFreezePanes) to clear all pane settings, then freezes panes at cell C3 and saves the workbook as ResetPanesDemo.xlsx.
    public static class WorksheetExtensions
    {
        public static void ResetPanes(this Worksheet ws)
        {
            ws.RemoveSplit();      // Remove any split window
            ws.UnFreezePanes();    // Unfreeze any frozen panes
        }
    }

    public class ResetPanesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Clear all existing splits and freezes
                sheet.ResetPanes();

                // Apply a new pane configuration (freeze panes at C3)
                sheet.FreezePanes("C3", 3, 3);

                // Save the workbook
                string outputPath = "ResetPanesDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ResetPanesDemo.Run();
        }
    }
}
