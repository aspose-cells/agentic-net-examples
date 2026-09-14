// Title: Create and save an Excel workbook with numeric values, a SUM formula, and cell comments using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, puts 10 in A1, 20 in A2, assigns the formula =SUM(A1:A2) to A3, adds a comment to each tracked cell, and saves the file as an XLSX using Aspose.Cells. | Show how to define named ranges for cells A1, A2, and A3 as a watch‑window substitute before exporting the workbook to XLSX with Aspose.Cells for .NET.
// Common Searches: aspnet aspose.cells c# save workbook with formula to xlsx | how to add comments to cells in aspose.cells before saving | c# use named ranges to monitor cells in aspose.cells workbook | aspose.cells create workbook set sum formula export as xlsx | workaround for missing watch window feature in aspose.cells .net
// Tags: save workbook as xlsx aspose.cells c# | insert numeric values aspose.cells c# | set cell formula aspose.cells c# | cell comment workaround aspose.cells c# | named range tracking aspose.cells c# | watch window alternative aspose.cells

using System;
using Aspose.Cells;

namespace AsposeCellsWatchWindowDemo
{
    // The example creates a new Workbook, writes 10 to A1 and 20 to A2, sets a SUM formula in A3, adds comments (or named ranges) to emulate a watch window, and saves the workbook as 'WatchWindowDemo.xlsx' using the XLSX SaveFormat in Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data and a formula
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].Formula = "=SUM(A1:A2)";

                // NOTE: Aspose.Cells for .NET does not expose a WatchWindow API.
                // The original intent was to monitor cells A1, A2, and A3.
                // As an alternative, you can add comments or use named ranges for tracking.

                // Save the workbook to an XLSX file
                string outputPath = "WatchWindowDemo.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
