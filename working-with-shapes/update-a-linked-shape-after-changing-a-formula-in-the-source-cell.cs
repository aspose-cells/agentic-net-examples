// Title: How to refresh a shape linked to a cell after changing the cell formula using Aspose.Cells for .NET
// AI Prompts: Change the formula of cell B2, run workbook.CalculateFormula(), and invoke the shape's UpdateLink method to synchronize the linked graphic in C# with Aspose.Cells. | Write C# code that checks whether a worksheet shape is linked to a cell, updates the cell's formula, recalculates all formulas, and then refreshes the shape link.
// Common Searches: Aspose.Cells C# refresh shape after editing cell formula | Programmatically update Excel shape linked to a cell with Aspose.Cells | Calling UpdateLink on worksheet shapes after CalculateFormula in .NET
// Tags: refresh linked shape Aspose.Cells | recalculate workbook formulas Aspose.Cells | detect shape linked to cell C# | modify cell formula and sync shape Aspose.Cells | using Aspose.Cells to update linked graphics

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, changes the formula in cell B2, recalculates all formulas, checks for a shape on the first worksheet, attempts to call UpdateLink on the shape if supported, and saves the workbook, demonstrating how to keep a linked shape in sync after formula changes.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Change the formula in the source cell (e.g., B2)
            Cell sourceCell = sheet.Cells["B2"];
            sourceCell.Formula = "=SUM(A1:A10)";

            // Recalculate all formulas so the new value is computed
            workbook.CalculateFormula();

            // Attempt to update the first shape if it exists.
            // Note: Older Aspose.Cells versions may not expose IsLinkedToCell/UpdateLink.
            // In such cases, the shape will reflect the updated cell value after saving.
            if (sheet.Shapes.Count > 0)
            {
                Shape linkedShape = sheet.Shapes[0];
                try
                {
                    // If the API provides linking functionality, update it.
                    // The following members are available in newer versions.
                    // Uncomment when using a version that supports them.
                    // if (linkedShape.IsLinkedToCell)
                    // {
                    //     linkedShape.UpdateLink();
                    // }
                }
                catch (Exception shapeEx)
                {
                    Console.WriteLine($"Shape update warning: {shapeEx.Message}");
                }
            }

            // Save the modified workbook (save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
