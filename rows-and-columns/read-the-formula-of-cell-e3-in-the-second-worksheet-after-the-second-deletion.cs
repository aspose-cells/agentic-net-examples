// Title: C# – Retrieve the formula of cell E3 in the second worksheet after deleting two rows with Aspose.Cells
// Description: This example loads an existing workbook, confirms a second worksheet is present, removes the first row and the subsequent row on that sheet, saves the changes, and then reads the Formula property of cell E3 using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# read cell formula | delete rows Aspose.Cells | second worksheet formula retrieval | E3 formula after row removal | Aspose.Cells .NET example
// Common Searches: how to get a cell formula after deleting rows in Aspose.Cells | C# Aspose.Cells delete first two rows then read E3 | retrieve updated formula from second sheet with Aspose.Cells | Aspose.Cells example for row deletion and formula extraction
// Developer Intent: Extract the current formula string of cell E3 on the workbook's second sheet after two rows have been removed.
// Use Cases: Validate that formulas shift correctly when rows are removed from a financial model. | Log or audit formulas in a template after structural edits. | Programmatically recalculate dependent values by reading adjusted formulas post‑modification.
// AI Prompts: Generate C# code that deletes the first two rows of the second worksheet and prints the formula of cell E3 using Aspose.Cells. | Explain how row deletions impact cell references in formulas and how Aspose.Cells reflects those changes. | Add robust error handling for missing worksheets, absent formulas, and file‑access issues in the provided Aspose.Cells snippet.

using System;
using System.IO;
using Aspose.Cells;

// This example loads an existing workbook, confirms a second worksheet is present, removes the first row and the subsequent row on that sheet, saves the changes, and then reads the Formula property of cell E3 using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Path to the original workbook
        string inputPath = "input.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure there is a second worksheet
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Reference to the second worksheet (index 1)
            Worksheet sheet = workbook.Worksheets[1];
            Cells cells = sheet.Cells;

            // First deletion: delete the first row (index 0)
            cells.DeleteRow(0);

            // Second deletion: delete the next row (now at index 1 after the first deletion)
            cells.DeleteRow(1);

            // Save the modified workbook to a temporary file
            string tempPath = "temp_modified.xlsx";
            workbook.Save(tempPath);

            // Retrieve the formula of cell E3 in the second worksheet
            // (Using direct cell access instead of CellsAI which may not be available)
            string formula = sheet.Cells["E3"].Formula;

            // Output the retrieved formula
            Console.WriteLine($"Formula in E3 of the second worksheet: {formula}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
