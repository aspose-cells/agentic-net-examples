// Title: Delete rows above a freeze pane and reset FreezePanes with Aspose.Cells for .NET
// AI Prompts: Delete the first N rows of a worksheet and then call FreezePanes(0,0,0,0) to keep the new top row frozen using Aspose.Cells. | After removing rows above the original freeze row, recalculate the freeze row index and apply the four‑parameter FreezePanes method to align the pane. | Programmatically shift a sheet upward by deleting rows and update the freeze pane with the FreezePanes overload in C#.
// Common Searches: Aspose.Cells how to delete rows above a freeze pane in C# | reset freeze pane after deleting top rows using Aspose.Cells .NET | C# example for adjusting FreezePanes coordinates after row removal | remove leading rows while keeping freeze pane in Excel with Aspose.Cells | Aspose.Cells FreezePanes overload row column totalRows totalColumns usage
// Tags: remove rows preceding freeze pane Aspose.Cells | reset FreezePanes coordinates after row removal | use FreezePanes(row, column, totalRows, totalColumns) Aspose.Cells | shift worksheet rows and maintain freeze pane .NET | Aspose.Cells top rows deletion with freeze pane update

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, deletes rows above a specified freeze row, recalculates the new freeze row index, applies the four‑parameter FreezePanes method to reflect the shifted layout, and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------------------------
                // Define the original freeze point (row index, zero‑based)
                // Example: freeze at row 5 (i.e., rows 0‑4 are frozen)
                int originalFreezeRow = 5;

                // Delete all rows above the intended freeze point
                // This removes rows 0 … originalFreezeRow‑1
                worksheet.Cells.DeleteRows(0, originalFreezeRow);

                // After deletion the sheet has shifted up. The new freeze row index
                // is the original index minus the number of rows removed, which becomes 0
                int newFreezeRow = 0;

                // Adjust the freeze panes to reflect the new layout.
                // Use the 4‑parameter overload: FreezePanes(row, column, totalRows, totalColumns)
                worksheet.FreezePanes(newFreezeRow, 0, newFreezeRow, 0);
                // -------------------------------------------------------------------

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
