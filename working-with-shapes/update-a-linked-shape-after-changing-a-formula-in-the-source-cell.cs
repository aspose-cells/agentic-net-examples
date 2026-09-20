// Title: Programmatically refresh linked OLE objects after changing a cell formula using Aspose.Cells for .NET
// AI Prompts: Change the formula of a target cell, run workbook.CalculateFormula(), and call Aspose.Cells to refresh any linked OLE objects in the worksheet. | Show how to modify a cell's formula, recalculate the workbook, and attempt to synchronize linked shapes such as OLE objects or pictures with the new values in C#.
// Common Searches: how to refresh linked OLE objects after editing a cell formula with Aspose.Cells C# | Aspose.Cells recalculate formulas and update linked pictures programmatically | C# example for syncing OLE objects with changed cell values using Aspose.Cells | update linked shape when source cell formula changes Aspose.Cells .NET
// Tags: refresh linked OLE objects Aspose.Cells | recalculate workbook formulas Aspose.Cells | update linked picture after formula change C# | Aspose.Cells OLE object synchronization | modify cell formula and sync shapes .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// The example loads an existing workbook, changes the formula in cell B2, recalculates all formulas, iterates through OLE objects (noting that Aspose.Cells currently lacks a direct method to update linked shapes), and saves the modified file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Change the formula in cell B2
            Cell sourceCell = sheet.Cells["B2"];
            sourceCell.Formula = "=SUM(A1:A10)";

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Iterate through OLE objects (linked objects handling not required by current API)
            foreach (OleObject ole in sheet.OleObjects)
            {
                try
                {
                    // Placeholder for any OLE object processing if needed
                    // Current Aspose.Cells API does not expose direct link update methods
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to process OLE object: {ex.Message}");
                }
            }

            // Save the workbook with the updated content
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
