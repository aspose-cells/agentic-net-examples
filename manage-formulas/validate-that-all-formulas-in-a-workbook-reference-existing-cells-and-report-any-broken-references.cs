// Title: Detect and list broken #REF! formula references in an Excel workbook with Aspose.Cells for C#
// Common Searches: how to detect #REF! errors in an Excel file using Aspose.Cells C# | C# code to list cells with broken formula references in a workbook with Aspose.Cells | validate all formulas in an Excel workbook and find invalid references using Aspose.Cells .NET | Aspose.Cells calculate formulas and check for #REF! results in C# | enumerate cells containing #REF! after formula calculation with Aspose.Cells | report broken formula references in Excel using Aspose.Cells library
// Tags: broken formula reference detection Aspose.Cells | formula cell enumeration C# | pre‑calculate workbook formulas Aspose.Cells | #REF! error listing Excel .NET | Excel formula validation Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, forces a full formula calculation, then iterates through each worksheet and every cell that contains a formula. It checks the evaluated cell value for the "#REF!" error, collects the sheet name, cell address, and formula for any broken references, and outputs the list or reports that no broken references were found.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Force calculation of all formulas so that error values are up‑to‑date
            try
            {
                workbook.CalculateFormula();
            }
            catch (Exception calcEx)
            {
                Console.WriteLine($"Formula calculation error: {calcEx.Message}");
            }

            // List to hold information about broken formula references
            List<string> brokenFormulas = new List<string>();

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each cell that contains a formula
                foreach (Cell cell in sheet.Cells)
                {
                    if (cell.IsFormula)
                    {
                        // Check if the formula result is a #REF! error (string comparison)
                        if (cell.Value != null && cell.Value.ToString() == "#REF!")
                        {
                            brokenFormulas.Add($"{sheet.Name}!{cell.Name}: {cell.Formula}");
                        }
                    }
                }
            }

            // Output the results
            if (brokenFormulas.Count > 0)
            {
                Console.WriteLine("Broken formula references found:");
                foreach (string info in brokenFormulas)
                {
                    Console.WriteLine(info);
                }
            }
            else
            {
                Console.WriteLine("No broken formula references detected.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
