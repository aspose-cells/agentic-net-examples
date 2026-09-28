// Title: C# Aspose.Cells utility to detect cells with QuotePrefix enabled in an Excel workbook
// AI Prompts: Write a C# method that loads an .xlsx file with Aspose.Cells, iterates over every worksheet's used range, and returns a list of A1‑style addresses where Style.QuotePrefix is true. | Enhance the scanner to also output each cell's value together with its address for all cells that have QuotePrefix set. | Create a batch version that accepts a folder path, processes every Excel file with Aspose.Cells, and writes the QuotePrefix‑enabled cell locations to a log file.
// Common Searches: how to find cells with QuotePrefix true using Aspose.Cells C# | list all Excel cells that have quote prefix flag set in .NET | detect unintended quote prefix style in a workbook programmatically | Aspose.Cells scan workbook for QuotePrefix property | C# script to report cells with QuotePrefix enabled in multiple Excel files
// Tags: Aspose.Cells detect QuotePrefix flag | C# scan Excel workbook for style property | list cells with QuotePrefix true | iterate used range Aspose.Cells | batch process Excel files for QuotePrefix detection

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    // Utility that loads a workbook with Aspose.Cells, walks through each worksheet's used rows and columns, checks the Style.QuotePrefix flag for every cell, collects the A1 addresses of cells where the flag is true, and writes the results to the console (or a log file in batch mode).
    public static class QuotePrefixScanner
    {
        /// <param name="inputPath">Path to the workbook to be scanned.</param>
        public static void Scan(string inputPath)
        {
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                Console.WriteLine("Input path is null or empty.");
                return;
            }

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (lifecycle rule: use provided load logic)
                Workbook workbook = new Workbook(inputPath);

                // List to hold addresses of cells with QuotePrefix set
                List<string> cellsWithQuotePrefix = new List<string>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    Cells cells = worksheet.Cells;

                    // Determine the used range to limit iteration
                    int maxRow = cells.MaxDataRow;
                    int maxColumn = cells.MaxDataColumn;

                    // Scan each cell within the used range
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxColumn; col++)
                        {
                            Cell cell = cells[row, col];

                            // Retrieve the cell's style and check the QuotePrefix property
                            Style style = cell.GetStyle();
                            if (style.QuotePrefix)
                            {
                                // Record the address in A1 notation with worksheet name
                                string address = $"{worksheet.Name}!{cell.Name}";
                                cellsWithQuotePrefix.Add(address);
                            }
                        }
                    }
                }

                // Output the results
                Console.WriteLine("Cells with QuotePrefix = true:");
                if (cellsWithQuotePrefix.Count == 0)
                {
                    Console.WriteLine("None found.");
                }
                else
                {
                    foreach (string addr in cellsWithQuotePrefix)
                    {
                        Console.WriteLine(addr);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while scanning the workbook: {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be inspected
            string workbookPath = "input.xlsx";

            // Run the scanner
            QuotePrefixScanner.Scan(workbookPath);
        }
    }
}
