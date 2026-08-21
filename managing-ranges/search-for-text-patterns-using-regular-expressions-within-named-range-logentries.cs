// Title: C# – Find Regex Patterns in the Named Range "LogEntries" with Aspose.Cells
// Description: Loads an Excel file, retrieves the named range "LogEntries", builds a matching CellArea, configures FindOptions for regular‑expression search, iterates over all cells that match a date pattern (YYYY‑MM‑DD), outputs their addresses, and saves the workbook.
// Keywords: Aspose.Cells regex search | C# find cells by pattern | named range FindOptions | Excel regex lookup | CellArea range search
// Common Searches: Aspose.Cells find regex in named range | C# search Excel cells with regular expression | Limit Aspose.Cells Find to a specific range | How to use FindOptions.RegexKey in .NET | Extract dates from a named range using Aspose
// Developer Intent: Retrieve every cell whose value matches a regular expression inside the "LogEntries" named range.
// Use Cases: Extract all date strings from a log sheet for reporting. | Validate that entries in a named range follow a required format. | Flag or highlight cells that meet a pattern before exporting data.
// AI Prompts: Write C# code that uses Aspose.Cells to locate email addresses in a named range "Contacts" and apply a yellow background. | Show how to configure FindOptions to search formulas with a regex that matches cell references within a specific range. | Provide an example that iterates over regex matches in a named range and copies the matched values to a new worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRegexSearch
{
    // Loads an Excel file, retrieves the named range "LogEntries", builds a matching CellArea, configures FindOptions for regular‑expression search, iterates over all cells that match a date pattern (YYYY‑MM‑DD), outputs their addresses, and saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the named range "LogEntries"
                Name logRangeName = workbook.Worksheets.Names["LogEntries"];
                if (logRangeName == null)
                {
                    Console.WriteLine("Named range 'LogEntries' not found.");
                    return;
                }

                // Obtain the Aspose.Cells.Range object from the named range
                Aspose.Cells.Range logRange = logRangeName.GetRange();

                // Build a CellArea that represents the same range
                int firstRow = logRange.FirstRow;
                int firstColumn = logRange.FirstColumn;
                int lastRow = firstRow + logRange.RowCount - 1;
                int lastColumn = firstColumn + logRange.ColumnCount - 1;

                CellArea searchArea = new CellArea
                {
                    StartRow = firstRow,
                    StartColumn = firstColumn,
                    EndRow = lastRow,
                    EndColumn = lastColumn
                };

                // Configure find options for regex search
                FindOptions findOptions = new FindOptions
                {
                    RegexKey = true,                         // Enable regular expression matching
                    LookInType = LookInType.Values,          // Search in cell values
                    LookAtType = LookAtType.EntireContent    // Exact match of the whole cell content
                };
                findOptions.SetRange(searchArea);            // Limit the search to the named range

                // Define the regular expression pattern to search for (e.g., dates YYYY-MM-DD)
                string regexPattern = @"\d{4}-\d{2}-\d{2}";

                // Perform the first search
                Cell foundCell = worksheet.Cells.Find(regexPattern, null, findOptions);

                // Iterate through all matching cells within the named range
                while (foundCell != null)
                {
                    Console.WriteLine($"Found match at {foundCell.Name}: {foundCell.StringValue}");

                    // Continue searching from the cell after the current one
                    foundCell = worksheet.Cells.Find(regexPattern, foundCell, findOptions);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
