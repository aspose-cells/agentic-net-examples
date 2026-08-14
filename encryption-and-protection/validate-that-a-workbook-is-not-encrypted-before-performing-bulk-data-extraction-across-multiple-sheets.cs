// Title: C# – Verify Workbook Encryption Before Bulk Extraction with Aspose.Cells for .NET
// Description: A concise C# example that loads an Excel file using Aspose.Cells, confirms the file exists, checks workbook.Settings.IsEncrypted, skips encrypted workbooks, and iterates every worksheet to read each used cell, outputting address and value with robust error handling.
// Keywords: Aspose.Cells | C# | .NET | workbook encryption | IsEncrypted | bulk data extraction | multiple worksheets | read cell values | Excel password protection | batch processing | exception handling | GitHub example
// Common Searches: How to check if an Excel workbook is encrypted with Aspose.Cells .NET | Skip password‑protected Excel files during bulk extraction using Aspose.Cells | Read all cells from every sheet after confirming workbook is not encrypted | Aspose.Cells C# example for validating encryption before processing | Batch import Excel files with encryption detection in .NET
// Developer Intent: Validate that a workbook is not encrypted before extracting data from all its worksheets.
// Use Cases: Prevent runtime errors in a batch import pipeline by ignoring password‑protected Excel files. | Log or migrate data from every sheet only when the workbook is confirmed unencrypted. | Integrate encryption checks into automated data‑migration or ETL processes that handle many Excel documents.
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells, returns true if workbook.Settings.IsEncrypted is false, and logs the result. | Create a method to extract all cell values from each worksheet of an unencrypted workbook and store them in a DataTable using Aspose.Cells. | Provide best‑practice error handling for processing a folder of Excel files where some may be encrypted, using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // A concise C# example that loads an Excel file using Aspose.Cells, confirms the file exists, checks workbook.Settings.IsEncrypted, skips encrypted workbooks, and iterates every worksheet to read each used cell, outputting address and value with robust error handling.
    public class BulkDataExtractionValidator
    {
        public static void Run(string filePath)
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook without a password first
                using (Workbook workbook = new Workbook(filePath))
                {
                    // Check if the workbook is encrypted
                    if (workbook.Settings.IsEncrypted)
                    {
                        Console.WriteLine($"The workbook \"{filePath}\" is encrypted and cannot be processed without a password.");
                        return;
                    }

                    Console.WriteLine($"Extracting data from workbook \"{filePath}\"...");

                    // Iterate through all worksheets
                    foreach (Worksheet sheet in workbook.Worksheets)
                    {
                        Console.WriteLine($"Worksheet: {sheet.Name}");

                        var cells = sheet.Cells;
                        int maxRow = cells.MaxDataRow;
                        int maxCol = cells.MaxDataColumn;

                        // Iterate through each cell in the used range
                        for (int row = 0; row <= maxRow; row++)
                        {
                            for (int col = 0; col <= maxCol; col++)
                            {
                                var cell = cells[row, col];
                                if (cell.Value != null)
                                {
                                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
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
                Console.Write("Enter the full path to the Excel file: ");
                filePath = Console.ReadLine();
            }

            BulkDataExtractionValidator.Run(filePath);
        }
    }
}
