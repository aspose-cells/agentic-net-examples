// Title: Extract formulas from a cell range with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, set sample formulas, define a range (A1:C3), loop through each cell, detect formulas, retrieve the A1‑notation text via Cell.GetFormula(false,false), collect address‑formula pairs, display them, and save the workbook—perfect for bulk formula analysis.
// Keywords: Aspose.Cells | C# | GetFormula | FormulaText extraction | extract formulas from range | bulk formula analysis | A1 notation | worksheet formulas | Aspose.Cells example | retrieve formula strings
// Common Searches: Aspose.Cells get formula text from range C# | How to extract all formulas in a worksheet using Aspose.Cells | Cell.GetFormula example for multiple cells | Bulk retrieve formulas Aspose.Cells .NET | Iterate over range and read formulas Aspose
// Developer Intent: Obtain the textual representation of every formula inside a specified cell range for auditing, reporting, or further processing.
// Use Cases: Audit all formulas in a financial model to verify calculation logic. | Export worksheet formulas to documentation or change‑log reports. | Compare original and transformed formulas after applying workbook macros or conversions.
// AI Prompts: Generate C# code that extracts formulas from any given Aspose.Cells range and returns a dictionary keyed by cell address. | Provide a robust example that logs each cell's address and formula text while skipping non‑formula cells. | Create a script that saves extracted formulas to a CSV file for downstream analysis.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Shows how to create a workbook, set sample formulas, define a range (A1:C3), loop through each cell, detect formulas, retrieve the A1‑notation text via Cell.GetFormula(false,false), collect address‑formula pairs, display them, and save the workbook—perfect for bulk formula analysis.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some sample formulas for demonstration
            cells["A1"].Formula = "=SUM(1,2)";                 // 3
            cells["B2"].Formula = "=A1*5";                    // 15
            cells["C3"].Formula = "=IF(A1>0,\"Yes\",\"No\")"; // Yes

            // Define the range whose formulas we want to extract
            Aspose.Cells.Range targetRange = cells.CreateRange("A1:C3");

            // Collection to hold extracted formula texts
            List<string> extractedFormulas = new List<string>();

            // Iterate through each cell in the range
            foreach (Cell cell in targetRange)
            {
                // Check if the cell actually contains a formula
                if (cell.IsFormula)
                {
                    // Get the formula text in A1 notation (non‑R1C1, non‑local)
                    string formulaText = cell.GetFormula(false, false);
                    extractedFormulas.Add($"{cell.Name}: {formulaText}");
                }
            }

            // Output the extracted formulas for bulk analysis
            Console.WriteLine("Extracted formulas from the range:");
            foreach (string entry in extractedFormulas)
            {
                Console.WriteLine(entry);
            }

            // Save the workbook (lifecycle save)
            workbook.Save("FormulaTextExtraction.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
