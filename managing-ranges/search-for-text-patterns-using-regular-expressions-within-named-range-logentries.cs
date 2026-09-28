// Title: Find cells matching a regular expression inside the named range 'LogEntries' using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to scan a named range and return the addresses of cells whose text matches a given regex pattern. | Show how to apply conditional formatting with Aspose.Cells to highlight cells in the 'LogEntries' range that contain error codes matching \bERROR\d{3}\b. | Create a reusable method that accepts a Workbook, a named range name, and a regex pattern, then logs each matching cell name and value.
// Common Searches: aspnet cells search regex within a specific named range example | c# Aspose.Cells find error codes in Excel named range | how to iterate over cells of a named range and apply Regex.IsMatch in Aspose.Cells | retrieve cell addresses that match pattern in Excel using Aspose.Cells .NET | using Aspose.Cells to locate text matching \bERROR\d{3}\b in a workbook
// Tags: regex search Aspose.Cells named range | iterate cells Aspose.Cells workbook | highlight matching cells Aspose.Cells | error code detection Excel .NET | named range boundary calculation Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel file, obtains the named range "LogEntries", iterates through each cell in that range, and uses a case‑insensitive regular expression (\bERROR\d{3}\b) to identify matching text, printing each match and then saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Retrieve the named range "LogEntries"
            AsposeRange namedRange = workbook.Worksheets.GetRangeByName("LogEntries");
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'LogEntries' not found.");
                return;
            }

            // Regular expression pattern to search for
            const string regexPattern = @"\bERROR\d{3}\b";

            // Determine range boundaries using Aspose.Cells properties
            int startRow = namedRange.FirstRow;
            int endRow = startRow + namedRange.RowCount - 1;
            int startCol = namedRange.FirstColumn;
            int endCol = startCol + namedRange.ColumnCount - 1;

            // Search cells within the named range and apply regex manually
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    Cell cell = namedRange.Worksheet.Cells[row, col];
                    string cellText = cell.StringValue ?? string.Empty;

                    if (Regex.IsMatch(cellText, regexPattern, RegexOptions.IgnoreCase))
                    {
                        Console.WriteLine($"Match found at {cell.Name}: {cellText}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (if any modifications were made)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
