// Title: Validate that an Excel workbook is not password‑protected before bulk extracting data from all worksheets using Aspose.Cells in C#
// Common Searches: how to detect if an Excel file is password protected using Aspose.Cells C# | bulk read all worksheets after confirming workbook is not encrypted with Aspose.Cells | C# code to verify Excel workbook encryption before iterating sheets | Aspose.Cells load workbook without password exception for multiple sheet extraction | extract data from every sheet in an unencrypted Excel workbook using Aspose.Cells
// Tags: detect encrypted workbook Aspose.Cells | bulk extract worksheets Aspose.Cells | handle password‑protected Excel Aspose.Cells | verify workbook encryption C# | iterate worksheets Aspose.Cells | use MaxDisplayRange Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace BulkDataExtraction
{
    // The example loads an Excel file with Aspose.Cells, automatically throws an exception if the file is missing or password‑protected, then iterates each worksheet, determines its used range via MaxDisplayRange, and copies every cell's string value into a dictionary keyed by sheet name, returning a list‑of‑lists representation while providing comprehensive error handling.
    public class WorkbookExtractor
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>
        /// A dictionary where the key is the worksheet name and the value is a list of rows,
        /// each row represented as a list of string values.
        /// </returns>
        public static Dictionary<string, List<List<string>>> ExtractData(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("The specified Excel file was not found.", filePath);

            try
            {
                // Load the workbook. If the file is password‑protected, an exception will be thrown.
                Workbook workbook = new Workbook(filePath);

                var result = new Dictionary<string, List<List<string>>>();

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    var sheetData = new List<List<string>>();

                    // Determine the used range of the sheet.
                    AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                    // If the sheet is empty, skip it.
                    if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                    {
                        result.Add(sheet.Name, sheetData);
                        continue;
                    }

                    int startRow = usedRange.FirstRow;
                    int endRow = startRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = startCol + usedRange.ColumnCount - 1;

                    // Loop through rows and columns within the used range.
                    for (int row = startRow; row <= endRow; row++)
                    {
                        var rowData = new List<string>();
                        for (int col = startCol; col <= endCol; col++)
                        {
                            // Retrieve the cell value as a string (null‑safe).
                            string cellValue = sheet.Cells[row, col].StringValue ?? string.Empty;
                            rowData.Add(cellValue);
                        }
                        sheetData.Add(rowData);
                    }

                    result.Add(sheet.Name, sheetData);
                }

                return result;
            }
            catch (Exception ex)
            {
                // Wrap any exception to provide context.
                throw new InvalidOperationException($"Failed to extract data from workbook '{filePath}'.", ex);
            }
        }

        // Entry point for the console application.
        public static void Main(string[] args)
        {
            // Example usage: provide the Excel file path as the first argument.
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the full path to the Excel file as a command‑line argument.");
                return;
            }

            string filePath = args[0];

            try
            {
                var data = ExtractData(filePath);
                Console.WriteLine($"Extracted data from '{Path.GetFileName(filePath)}':");

                foreach (var sheetEntry in data)
                {
                    Console.WriteLine($"\nWorksheet: {sheetEntry.Key}");
                    foreach (var row in sheetEntry.Value)
                    {
                        Console.WriteLine(string.Join("\t", row));
                    }
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine(fnfEx.Message);
            }
            catch (InvalidOperationException invEx)
            {
                Console.WriteLine(invEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
