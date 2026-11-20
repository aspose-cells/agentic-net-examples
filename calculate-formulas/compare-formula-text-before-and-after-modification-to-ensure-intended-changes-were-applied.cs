// Title: Compare an Excel cell's original and modified formula text using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# compare original and updated cell formula | detect changes in Excel formula after edit with Aspose.Cells | retrieve cell formula before and after modification in .NET | how to verify if a formula was altered using Aspose.Cells | save workbook after changing formula and check difference Aspose.Cells
// Tags: Aspose.Cells get cell formula text | C# check formula equality | Aspose.Cells replace operator in formula | Aspose.Cells save workbook after formula edit | detect formula modification .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook (or creates a new one), reads the formula from cell B2, optionally replaces '+' with '-', assigns a new formula if none exists, captures the updated formula, compares the original and new formula strings for equality, reports whether the formula changed, and saves the workbook to a new file.
class FormulaComparison
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        Workbook workbook = null;

        try
        {
            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                // Attempt to load the workbook (no password)
                workbook = new Workbook(inputPath);
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";

                // Optionally add a sample formula to B2 for demonstration
                workbook.Worksheets[0].Cells["B2"].Formula = "=A1+A2";
            }
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Work with the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Specify the cell to work with (e.g., B2)
        Cell cell = worksheet.Cells["B2"];

        // Capture the original formula text (if any)
        string originalFormula = cell.Formula;

        // Display the original formula
        Console.WriteLine($"Original formula in {cell.Name}: {originalFormula}");

        // Modify the formula (example: change addition to subtraction)
        if (!string.IsNullOrEmpty(originalFormula))
        {
            // Simple replacement logic for demonstration
            string modifiedFormula = originalFormula.Replace("+", "-");
            cell.Formula = modifiedFormula;
        }
        else
        {
            // If the cell had no formula, assign a new one
            cell.Formula = "=SUM(A1:A5)";
        }

        // Capture the formula after modification
        string newFormula = cell.Formula;

        // Display the new formula
        Console.WriteLine($"Modified formula in {cell.Name}: {newFormula}");

        // Compare the original and new formula texts
        bool formulasChanged = !string.Equals(originalFormula, newFormula, StringComparison.Ordinal);
        Console.WriteLine($"Formula changed: {formulasChanged}");

        // Save the workbook to persist changes (optional)
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
