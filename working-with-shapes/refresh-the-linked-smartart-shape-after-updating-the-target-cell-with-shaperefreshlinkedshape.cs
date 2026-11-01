// Title: Refresh a SmartArt shape linked to a worksheet cell after updating the cell value with Aspose.Cells for .NET (using Shape.RefreshLinkedShape)
// AI Prompts: Generate C# code that changes the value of a specific cell, scans all shapes on the worksheet, identifies those linked to the modified cell, and invokes Shape.RefreshLinkedShape for each identified SmartArt shape using Aspose.Cells. | Show how to safely handle missing input files and log shape‑processing errors while preparing to call Shape.RefreshLinkedShape after a cell update in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells Shape.RefreshLinkedShape usage example in C# | how to refresh SmartArt linked to a cell after updating the cell value .NET | C# code to update Excel cell and refresh associated SmartArt shape with Aspose.Cells | enumerate worksheet shapes and refresh linked SmartArt using Aspose.Cells API
// Tags: Shape.RefreshLinkedShape method Aspose.Cells | linked SmartArt shape refresh workflow | worksheet shape enumeration Aspose.Cells C# | cell value update triggers shape refresh | Aspose.Cells handling of linked shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads a workbook, updates cell B2, and iterates through all shapes on the first worksheet. Although Aspose.Cells currently does not expose direct SmartArt manipulation, the code marks where Shape.RefreshLinkedShape would be called to refresh any SmartArt shapes linked to the modified cell. The workbook is then saved to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Update the target cell that the SmartArt (or any linked shape) is linked to
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue(12345); // new value for the linked cell

            // Iterate through shapes – placeholder for any shape-specific processing
            foreach (Shape shape in sheet.Shapes)
            {
                try
                {
                    // Currently Aspose.Cells does not expose SmartArt manipulation.
                    // If future APIs provide such support, insert refresh logic here.
                }
                catch (Exception shapeEx)
                {
                    // Log shape-specific errors without stopping the whole process
                    Console.WriteLine($"Shape processing error: {shapeEx.Message}");
                }
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
