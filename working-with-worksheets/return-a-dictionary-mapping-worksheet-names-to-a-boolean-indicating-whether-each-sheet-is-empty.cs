// Title: How to detect empty worksheets in an Excel file and return a C# dictionary of sheet names using Aspose.Cells
// AI Prompts: Write a C# method that loads an Excel workbook with Aspose.Cells and returns a Dictionary<string, bool> where each key is a worksheet name and the value is true if the sheet contains no data. | Build a console program that takes a file path argument, calls the method to find empty sheets, and prints each worksheet name followed by "Empty" or "Not Empty".
// Common Searches: Aspose.Cells C# check if a worksheet is empty | Get list of empty sheets from Excel workbook using Aspose.Cells | C# dictionary of worksheet names to emptiness status Aspose.Cells | Determine empty worksheets by MaxDataRow and MaxDataColumn in Aspose.Cells | How to programmatically find blank worksheets in .xlsx with Aspose.Cells
// Tags: Aspose.Cells empty worksheet detection | C# identify blank sheets in Excel | MaxDataRow MaxDataColumn emptiness test | Excel workbook sheet emptiness mapping | dictionary of sheet name to empty flag

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace WorksheetEmptyCheckerApp
{
    // The example loads an Excel file with Aspose.Cells, iterates through each worksheet, uses MaxDataRow and MaxDataColumn to determine if a sheet has no data, and returns a Dictionary<string,bool> mapping worksheet names to true when the sheet is empty.
    public class WorksheetEmptyChecker
    {
        /// <param name="filePath">Full path to the Excel file.</param>
        /// <returns>Dictionary mapping worksheet names to a boolean (true if empty).</returns>
        public static Dictionary<string, bool> GetEmptySheets(string filePath)
        {
            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            Workbook workbook;
            try
            {
                // Load the workbook from the file
                workbook = new Workbook(filePath);
            }
            catch (Exception ex)
            {
                // Wrap any loading errors for clearer diagnostics
                throw new InvalidOperationException("Failed to load workbook.", ex);
            }

            // Prepare the result dictionary
            var sheetEmptyMap = new Dictionary<string, bool>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // A sheet is considered empty when there are no cells with data.
                // MaxDataRow and MaxDataColumn return -1 if the sheet contains no data.
                bool isEmpty = sheet.Cells.MaxDataRow == -1 && sheet.Cells.MaxDataColumn == -1;

                // Add the result to the dictionary using the sheet's name as the key
                sheetEmptyMap[sheet.Name] = isEmpty;
            }

            return sheetEmptyMap;
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            // Expect the first argument to be the Excel file path
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WorksheetEmptyCheckerApp <excel-file-path>");
                return;
            }

            string filePath = args[0];

            try
            {
                var emptySheets = WorksheetEmptyChecker.GetEmptySheets(filePath);
                foreach (var kvp in emptySheets)
                {
                    Console.WriteLine($"Worksheet \"{kvp.Key}\": {(kvp.Value ? "Empty" : "Not Empty")}");
                }
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"Error: {fnfEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
