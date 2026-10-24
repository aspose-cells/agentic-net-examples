// Title: C# batch script to open multiple Excel workbooks and set a shared SUM formula using Aspose.Cells
// Common Searches: how to apply the same shared formula to several Excel files with Aspose.Cells in C# | C# code for batch updating workbooks with a shared SUM(B1:B10) formula | using Aspose.Cells to set shared formulas across multiple .xlsx files programmatically | automate shared formula insertion in a list of Excel workbooks with .NET | process a collection of Excel workbooks and overwrite them after adding a shared formula
// Tags: Aspose.Cells set shared formula C# | batch process Excel workbooks Aspose.Cells | shared formula group index Aspose | programmatic Excel formula insertion .NET | overwrite original workbook after formula change | iterate workbook files Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The C# program iterates over a list of Excel file paths, loads each workbook with Aspose.Cells, defines a shared SUM(B1:B10) formula in cell A1, propagates the shared formula to cells A2‑A10, saves the workbook back to its original location, and logs missing files or processing errors.
class Program
{
    static void Main()
    {
        // List of workbook file paths to process
        List<string> workbookFiles = new List<string>
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            @"C:\Workbooks\Book3.xlsx"
            // Add more file paths as needed
        };

        foreach (string filePath in workbookFiles)
        {
            try
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Define the shared formula (sums column B for rows 1‑10)
                string sharedFormula = "SUM(B1:B10)";

                // Index for the shared formula group (0 = first group)
                int sharedFormulaIndex = 0;

                // Set the shared formula for the master cell (A1)
                // Third parameter is the shared formula group index (use 0 for master)
                cells["A1"].SetSharedFormula(sharedFormula, sharedFormulaIndex, 0);

                // Apply the shared formula to the rest of the range (A2:A10)
                for (int row = 1; row < 10; row++) // rows are zero‑based; row 1 = A2
                {
                    // Use the same shared formula index; group index remains 0
                    cells[row, 0].SetSharedFormula(null, sharedFormulaIndex, 0);
                }

                // Save the workbook (overwrites the original file)
                workbook.Save(filePath);
                Console.WriteLine($"Processed and saved: {filePath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors for this file and continue with the next one
                Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
            }
        }
    }
}
