// Title: C# utility to detect cells with QuotePrefix = true using Aspose.Cells
// Description: Loads an Excel workbook with Aspose.Cells, scans each worksheet's used range, checks the QuotePrefix flag in cell styles, collects the addresses of cells where the flag is true, and prints the list to the console.
// Keywords: Aspose.Cells QuotePrefix detection | C# scan Excel cells QuotePrefix | find cells with leading apostrophe Aspose | list cells QuotePrefix true .NET | Excel style QuotePrefix check
// Common Searches: how to find cells with QuotePrefix set in Aspose.Cells | C# code to list Excel cells where QuotePrefix is true | Aspose.Cells iterate used range and check QuotePrefix flag | detect unintended leading apostrophe in Excel with Aspose
// Developer Intent: Locate every cell in a workbook that has the QuotePrefix property enabled.
// Use Cases: Validate imported spreadsheets to ensure no accidental leading‑apostrophe formatting. | Generate a report of cells that may display a leading apostrophe for quality control. | Programmatically clear the QuotePrefix flag on identified cells before further processing.
// AI Prompts: Create a method that returns a list of cell addresses with QuotePrefix = true using Aspose.Cells for .NET. | Optimize the scanner to reduce style object creation while checking QuotePrefix. | Show how to reset the QuotePrefix flag for the detected cells and save the workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Loads an Excel workbook with Aspose.Cells, scans each worksheet's used range, checks the QuotePrefix flag in cell styles, collects the addresses of cells where the flag is true, and prints the list to the console.
class QuotePrefixScanner
{
    static void Main(string[] args)
    {
        // Path to the workbook to be scanned
        string workbookPath = "input.xlsx";

        // Load the workbook (uses Aspose.Cells built‑in load logic)
        Workbook workbook = new Workbook(workbookPath);

        // List to hold addresses of cells where QuotePrefix is true
        List<string> cellsWithQuotePrefix = new List<string>();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Scan only the used range to improve performance
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Get the cell's style and check the QuotePrefix flag
                    if (cell != null && cell.GetStyle().QuotePrefix)
                    {
                        // Record the cell address with sheet name (e.g., Sheet1!B2)
                        cellsWithQuotePrefix.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }
        }

        // Output the results
        Console.WriteLine("Cells with QuotePrefix set to true:");
        foreach (string address in cellsWithQuotePrefix)
        {
            Console.WriteLine(address);
        }
    }
}
