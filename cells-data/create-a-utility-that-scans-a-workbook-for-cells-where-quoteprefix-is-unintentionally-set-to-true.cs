// Title: C# utility to scan an Excel workbook and list cells where the QuotePrefix style flag is true using Aspose.Cells
// Common Searches: C# code to find all cells with QuotePrefix enabled in an Excel file using Aspose.Cells | how to list cell addresses where QuotePrefix is true in a .NET workbook | detect unintentionally set QuotePrefix property in Excel worksheets with Aspose.Cells | enumerate quoted cells in an Excel workbook programmatically in C# | Aspose.Cells scan workbook for cells that have QuotePrefix flag set
// Tags: quote prefix detection Aspose.Cells | excel cell style scanning C# | list cells with QuotePrefix flag .NET | workbook style inspection Aspose.Cells | enumerate quoted cells in .xlsx using C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace QuotePrefixScanner
{
    // The example loads an Excel workbook with Aspose.Cells, iterates through each worksheet and its non‑empty cells, checks the Style.QuotePrefix flag for each cell, records the addresses of cells where the flag is true, and prints a list of those cell addresses.
    class Program
    {
        /// <param name="args">
        /// args[0] - Path to the input Excel file.
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: QuotePrefixScanner <input-file-path>");
                return;
            }

            string inputPath = args[0];

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(inputPath);

            // List to hold addresses of cells with QuotePrefix = true
            List<string> quotedCells = new List<string>();

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the Cells collection for the current worksheet
                Cells cells = sheet.Cells;

                // Enumerate only the cells that contain data (non-empty cells)
                foreach (Cell cell in cells)
                {
                    // Retrieve the style of the current cell
                    Style style = cell.GetStyle();

                    // Check if QuotePrefix is set to true
                    if (style.QuotePrefix)
                    {
                        // Record the cell address in the format SheetName!A1
                        quotedCells.Add($"{sheet.Name}!{cell.Name}");
                    }
                }
            }

            // Output the results
            if (quotedCells.Count == 0)
            {
                Console.WriteLine("No cells with QuotePrefix set to true were found.");
            }
            else
            {
                Console.WriteLine("Cells with QuotePrefix set to true:");
                foreach (string address in quotedCells)
                {
                    Console.WriteLine(address);
                }
            }
        }
    }
}
