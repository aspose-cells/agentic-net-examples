// Title: C# example: Load an Excel workbook with Aspose.Cells and verify that specific cells contain data (simulating the Watch Window)
// Common Searches: C# Aspose.Cells load workbook and check if cell A1 has a value | how to programmatically verify multiple cell addresses in an Excel file using Aspose.Cells | Aspose.Cells C# sample to iterate over a list of cells and report empty or populated | validate presence of cells in a saved .xlsx file with Aspose.Cells .NET | simulate Excel Watch Window verification with Aspose.Cells code
// Tags: Aspose.Cells load workbook from .xlsx | Aspose.Cells read cell value C# | Aspose.Cells iterate cell addresses | Aspose.Cells verify non‑empty cells | Aspose.Cells workbook cell existence check

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The C# program uses Aspose.Cells to open an existing Excel file (output.xlsx), defines a list of cell addresses (e.g., A1, B2, C3), and iterates through them on the first worksheet. For each cell it checks whether the cell object exists and contains a non‑null value, outputting a message that simulates the verification of those cells' presence in Excel's Watch Window. The example includes file‑existence handling and exception management.
class WatchWindowVerifier
{
    static void Main()
    {
        // Path to the saved workbook
        string workbookPath = "output.xlsx";

        // Ensure the workbook file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(workbookPath);

            // Define the cell addresses that should appear in the Watch Window
            // Adjust these addresses as needed for your verification
            List<string> cellsToVerify = new List<string> { "A1", "B2", "C3" };

            // Access the first worksheet (or modify to select a specific one)
            Worksheet worksheet = workbook.Worksheets[0];

            // Verify each specified cell address by checking if the cell contains any value
            foreach (string cellAddress in cellsToVerify)
            {
                bool found = false;

                try
                {
                    // Retrieve the cell; Cells[] returns a Cell object even if empty
                    Cell cell = worksheet.Cells[cellAddress];
                    // Consider the cell "present" if it has a non‑null value
                    if (cell != null && cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        found = true;
                    }
                }
                catch (Exception ex)
                {
                    // Handle any unexpected errors while accessing the cell
                    Console.WriteLine($"Error accessing cell {cellAddress}: {ex.Message}");
                }

                // Output verification result
                if (found)
                    Console.WriteLine($"Cell {cellAddress} contains data (simulating presence in Watch Window).");
                else
                    Console.WriteLine($"Cell {cellAddress} is empty or not found (simulating absence in Watch Window).");
            }
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., loading errors)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
