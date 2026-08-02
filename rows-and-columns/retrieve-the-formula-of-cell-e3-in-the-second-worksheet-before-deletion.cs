// Title: C# – Retrieve the formula of cell E3 on the second worksheet with Aspose.Cells
// Description: Loads an existing Excel file, accesses the worksheet at index 1 (the second sheet), reads the Formula property of cell E3, outputs the formula, optionally clears the cell, and saves the workbook. Demonstrates how to capture a formula before any deletion occurs.
// Keywords: Aspose.Cells C# | read cell formula | Excel formula retrieval | second worksheet | cell E3 formula | Aspose.Cells get formula | C# Excel automation | clear cell value Aspose
// Common Searches: Aspose.Cells get formula of a cell in .NET | C# read formula from second sheet Excel | retrieve Excel cell formula before clearing | how to obtain cell E3 formula using Aspose.Cells | read and delete cell formula Aspose.Cells C#
// Developer Intent: Extract the formula string from cell E3 on the second worksheet before removing its contents.
// Use Cases: Audit or log formulas prior to bulk data cleanup | Backup formulas to a separate repository before transformation | Validate expected calculations exist before applying worksheet updates | Generate documentation of dynamic formulas in multi‑sheet workbooks
// AI Prompts: Generate C# code that returns the formula of cell E3 on the second worksheet using Aspose.Cells without modifying the file. | Create a reusable method accepting a file path, sheet index, and cell address, which returns the cell's formula while preserving the original workbook. | Write a script that prints the formula of a specified cell, clears its value and formula, and then saves the workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing Excel file, accesses the worksheet at index 1 (the second sheet), reads the Formula property of cell E3, outputs the formula, optionally clears the cell, and saves the workbook. Demonstrates how to capture a formula before any deletion occurs.
class RetrieveFormulaDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the second worksheet (index 1)
            Worksheet secondSheet = workbook.Worksheets[1];

            // Get the cell E3 from the second worksheet
            Cell targetCell = secondSheet.Cells["E3"];

            // Retrieve the formula before any deletion occurs
            string formula = targetCell.Formula;
            Console.WriteLine($"Formula in worksheet '{secondSheet.Name}' cell E3: {formula}");

            // Example deletion: clear the cell's contents (removes the formula)
            targetCell.PutValue(string.Empty); // removes value and formula

            // Save the workbook after the operation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
