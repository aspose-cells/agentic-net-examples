using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    /// <summary>
    /// Scans a workbook and reports cells that have the QuotePrefix style flag set to true.
    /// </summary>
    public static class QuotePrefixScanner
    {
        /// <summary>
        /// Loads the workbook from the specified file, scans all worksheets and cells,
        /// and writes the addresses of cells with QuotePrefix = true to the console.
        /// </summary>
        /// <param name="filePath">Path to the workbook to be scanned.</param>
        public static void Scan(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found – \"{filePath}\"");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // List to hold addresses of cells with QuotePrefix set
                List<string> cellsWithQuotePrefix = new List<string>();

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Cells cells = sheet.Cells;

                    // Iterate through all used cells in the worksheet
                    for (int row = 0; row <= cells.MaxDataRow; row++)
                    {
                        for (int col = 0; col <= cells.MaxDataColumn; col++)
                        {
                            Cell cell = cells[row, col];
                            Style style = cell.GetStyle();

                            // Record cells where QuotePrefix is true
                            if (style.QuotePrefix)
                            {
                                string address = $"{sheet.Name}!{cell.Name}";
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

    /// <summary>
    /// Entry point for the console application.
    /// </summary>
    public static class Program
    {
        public static void Main(string[] args)
        {
            string filePath;

            if (args.Length > 0)
            {
                filePath = args[0];
            }
            else
            {
                Console.Write("Enter the path to the workbook to scan: ");
                filePath = Console.ReadLine();
            }

            QuotePrefixScanner.Scan(filePath);
        }
    }
}