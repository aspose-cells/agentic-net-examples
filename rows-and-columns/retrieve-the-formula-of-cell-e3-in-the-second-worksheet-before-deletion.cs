// Title: Aspose.Cells C# – Retrieve formula of E3 in second worksheet before clearing
// Description: Loads an existing workbook, accesses worksheet index 1, reads the formula from cell E3, prints it, empties the cell, and saves the workbook. Includes file‑existence verification and robust exception handling.
// Keywords: Aspose.Cells | C# | read cell formula | E3 formula | second worksheet | clear cell value | Excel formula extraction | Workbook.Save | exception handling
// Common Searches: Aspose.Cells get formula from cell E3 | C# read Excel formula before deleting cell | How to retrieve and clear cell content with Aspose.Cells | Get formula of a cell in specific worksheet Aspose | Aspose.Cells example read formula then clear
// Developer Intent: Extract the formula stored in cell E3 of the second worksheet before removing its content.
// Use Cases: Log original formulas before performing bulk data cleanup. | Validate and audit formulas in a particular sheet prior to exporting a sanitized workbook. | Capture cell formulas for compliance reporting before programmatically clearing cells.
// AI Prompts: Show C# code using Aspose.Cells to read the formula of cell E3 on the second worksheet and then clear the cell while preserving the workbook. | Explain how to safely retrieve a cell's formula before modifying its contents with Aspose.Cells for .NET. | Provide guidance for handling cases where the target cell does not contain a formula when using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing workbook, accesses worksheet index 1, reads the formula from cell E3, prints it, empties the cell, and saves the workbook. Includes file‑existence verification and robust exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (index 1)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Get the cell at E3
            Cell targetCell = secondSheet.Cells["E3"];

            // Retrieve the formula before any deletion
            string formula = targetCell.Formula;
            Console.WriteLine("Formula in E3 before deletion: " + formula);

            // Clear the cell's contents (using PutValue with an empty string)
            targetCell.PutValue(string.Empty);

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
