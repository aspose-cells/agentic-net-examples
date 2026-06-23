using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsRegexSearch
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the named range "LogEntries" is on the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Name logName = workbook.Worksheets.Names["LogEntries"];

                if (logName == null)
                {
                    Console.WriteLine("Named range 'LogEntries' not found.");
                    return;
                }

                // Get the address of the named range (e.g., =Sheet1!$A$1:$C$10)
                string refAddress = logName.RefersTo;
                if (refAddress.StartsWith("="))
                    refAddress = refAddress.Substring(1); // remove leading '='

                // Create a Range object from the address
                Aspose.Cells.Range logRange = worksheet.Cells.CreateRange(refAddress);

                // Convert the Range to a CellArea for FindOptions
                CellArea searchArea = new CellArea
                {
                    StartRow = logRange.FirstRow,
                    StartColumn = logRange.FirstColumn,
                    EndRow = logRange.FirstRow + logRange.RowCount - 1,
                    EndColumn = logRange.FirstColumn + logRange.ColumnCount - 1
                };

                // Configure FindOptions for regex search within the specified range
                FindOptions findOptions = new FindOptions
                {
                    RegexKey = true,                     // treat the search key as a regular expression
                    LookInType = LookInType.Values,      // search in cell values
                    LookAtType = LookAtType.EntireContent // exact match of the whole cell content
                };
                findOptions.SetRange(searchArea);

                // Define the regular expression pattern to search for (e.g., dates in format YYYY-MM-DD)
                string regexPattern = @"\d{4}-\d{2}-\d{2}";

                // Perform the first search
                Cell foundCell = worksheet.Cells.Find(regexPattern, null, findOptions);

                // Iterate through all matches within the named range
                while (foundCell != null)
                {
                    Console.WriteLine($"Found match at {foundCell.Name}: {foundCell.StringValue}");

                    // Continue searching from the cell after the current one
                    foundCell = worksheet.Cells.Find(regexPattern, foundCell, findOptions);
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
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