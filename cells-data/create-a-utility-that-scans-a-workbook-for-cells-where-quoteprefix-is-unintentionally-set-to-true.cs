// Title: C# Aspose.Cells utility to detect cells with QuotePrefix enabled
// Description: A console utility that loads an Excel workbook with Aspose.Cells, iterates the used range of each worksheet, checks the Style.QuotePrefix flag for every cell, and outputs the addresses of cells where QuotePrefix is true.
// Keywords: Aspose.Cells | QuotePrefix | C# | .NET | Excel | cell style | detect QuotePrefix | scan workbook | leading apostrophe | style flag | used range
// Common Searches: Aspose.Cells find cells with QuotePrefix | C# scan Excel for QuotePrefix true | list cells where QuotePrefix is set in .NET | detect leading apostrophe using Aspose.Cells | how to check QuotePrefix property in Aspose.Cells
// Developer Intent: Locate and list all cells that have the QuotePrefix style flag set to true.
// Use Cases: Validate incoming spreadsheets to prevent unintended text formatting. | Identify cells that may cause incorrect CSV export because of QuotePrefix. | Generate a report for data cleaning before further processing. | Automate removal of QuotePrefix after detection. | Audit workbooks for consistency of cell styles.
// AI Prompts: Generate code to clear the QuotePrefix flag for the cells returned by the scanner and save the workbook. | Modify the utility to also output each cell's original value and its full style information. | Write unit tests that verify the scanner correctly identifies cells with QuotePrefix set to true. | Create a PowerShell wrapper that runs the QuotePrefix scanner on multiple files in a folder. | Provide a step‑by‑step guide for integrating the scanner into an existing data‑import pipeline.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace QuotePrefixScannerUtility
{
    // A console utility that loads an Excel workbook with Aspose.Cells, iterates the used range of each worksheet, checks the Style.QuotePrefix flag for every cell, and outputs the addresses of cells where QuotePrefix is true.
    public class QuotePrefixScanner
    {
        /// <param name="inputFilePath">Path to the workbook to be scanned.</param>
        public static void Scan(string inputFilePath)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"Error: The file \"{inputFilePath}\" was not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            var cellsWithQuotePrefix = new List<string>();

            try
            {
                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    Cells cells = worksheet.Cells;

                    // Determine the used range to avoid iterating over empty cells
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;

                    // Scan each cell within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Skip null cells
                            if (cell == null)
                                continue;

                            // Skip cells that contain no data
                            if (cell.Type == CellValueType.IsNull)
                                continue;

                            // Retrieve the cell's style
                            Style style = cell.GetStyle();

                            // Check if QuotePrefix is set to true
                            if (style.QuotePrefix)
                            {
                                // Record the cell address in "SheetName!A1" format
                                cellsWithQuotePrefix.Add($"{worksheet.Name}!{cell.Name}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during scanning: {ex.Message}");
                return;
            }

            // Output the results
            Console.WriteLine("Cells with QuotePrefix set to true:");
            if (cellsWithQuotePrefix.Count == 0)
            {
                Console.WriteLine("None found.");
            }
            else
            {
                foreach (string address in cellsWithQuotePrefix)
                {
                    Console.WriteLine(address);
                }
            }
        }

        // Example usage
        public static void Main(string[] args)
        {
            // Ensure an input path is provided
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the workbook as a command‑line argument.");
                return;
            }

            string workbookPath = args[0];
            Scan(workbookPath);
        }
    }
}
