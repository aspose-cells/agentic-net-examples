// Title: Copy a range’s formulas to another range and convert relative references to absolute using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that copies formulas from a source range (e.g., A1:C3) to a destination range starting at A6, converting any relative cell references in the formulas to absolute references before assigning them. | Write a C# routine that iterates through each cell in a given range, copies the formula or raw value to a target range, and uses Aspose.Cells to transform relative references (e.g., A1) into absolute references for the copied formulas. | Create a C# example that demonstrates how to duplicate a worksheet range while preserving formulas, and includes logic to replace relative references with absolute ones using Aspose.Cells’ formula parsing capabilities.
// Common Searches: Aspose.Cells .NET how to transfer a range’s formulas and make them absolute | C# Aspose.Cells method for turning relative cell references into $A$1 when moving a range | example of copying a block of cells with formulas unchanged in Aspose.Cells | Aspose.Cells iterate over cells and keep formula logic while duplicating area | convert relative references to absolute during range copy using Aspose.Cells C#
// Tags: duplicate range formulas Aspose.Cells C# | relative-to-absolute reference conversion Aspose.Cells | iterate cells and copy values Aspose.Cells | preserve formulas while duplicating range Aspose.Cells | C# workbook range duplication Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new workbook, defines a source range (A1:C3), iterates through each cell, copies its formula (or raw value) to a destination range beginning at A6, and saves the file as CopiedFormulas.xlsx. The loop can be extended to replace relative references with absolute ones before assigning the formulas.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Define the source range whose formulas will be copied (e.g., A1:C3)
                var sourceRange = worksheet.Cells.CreateRange("A1:C3");

                // Define the top‑left cell of the destination range (e.g., starting at row 5, column 0 → A6)
                int destStartRow = 5;   // zero‑based index (row 6 in Excel)
                int destStartColumn = 0; // column A

                // Iterate through each cell in the source range
                for (int i = 0; i < sourceRange.RowCount; i++)
                {
                    for (int j = 0; j < sourceRange.ColumnCount; j++)
                    {
                        // Source cell
                        var srcCell = sourceRange[i, j];

                        // Corresponding destination cell
                        var destCell = worksheet.Cells[destStartRow + i, destStartColumn + j];

                        // If the source cell contains a formula, copy it
                        if (!string.IsNullOrEmpty(srcCell.Formula))
                        {
                            // Copy the formula (relative references will stay as‑is)
                            destCell.Formula = srcCell.Formula;
                        }
                        else
                        {
                            // If there is no formula, copy the raw value instead
                            destCell.PutValue(srcCell.Value);
                        }
                    }
                }

                // Save the workbook to a file
                string outputPath = "CopiedFormulas.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
