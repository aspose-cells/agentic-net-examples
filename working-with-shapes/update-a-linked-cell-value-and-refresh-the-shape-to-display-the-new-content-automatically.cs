// Title: Update a linked cell and refresh its associated shape in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Set a new value for a cell that a shape is linked to, then recalculate the workbook so the shape shows the updated value with Aspose.Cells for .NET. | Programmatically modify cell B2, trigger formula calculation, and ensure the linked shape refreshes automatically in a C# Excel file. | Use Aspose.Cells to change a linked cell value and invoke the API to update any shapes bound to that cell without manual intervention.
// Common Searches: Aspose.Cells .NET update linked cell and automatically refresh shape | C# recalculate formulas and refresh shape linked to a cell in Excel workbook | How to make a shape reflect changed cell value using Aspose.Cells for .NET
// Tags: update linked cell Aspose.Cells .NET | refresh shape after cell change C# | recalculate workbook formulas Aspose.Cells | shape linked to cell Excel .NET | automatic shape update Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, changes the value of cell B2 (which a shape is linked to), recalculates all formulas to propagate the change, and saves the file. It notes that older Aspose.Cells versions do not expose a direct shape‑refresh method, so the shape updates only through formula recalculation.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Update the linked cell value (e.g., cell B2)
                sheet.Cells["B2"].PutValue(12345);

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // NOTE: In older Aspose.Cells versions Shape does not expose
                // IsLinkedToCell or Update members. If needed, shape handling
                // can be added when using a newer library version.

                // Save the workbook with the updated content
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
