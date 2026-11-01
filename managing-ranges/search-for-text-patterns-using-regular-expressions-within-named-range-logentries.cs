// Title: Use Aspose.Cells for .NET to find and highlight cells matching a regex in the named range "LogEntries"
// Common Searches: aspnet cells find regex pattern in named range LogEntries using Aspose.Cells | c# highlight Excel cells that match ^ERROR: pattern with Aspose.Cells | how to apply conditional formatting based on regex in Aspose.Cells workbook | search and style string cells with regular expression in Aspose.Cells C# | retrieve named range and color cells matching regex Aspose.Cells .NET
// Tags: regex cell search Aspose.Cells | style cells based on regex Aspose.Cells | named range LogEntries Aspose.Cells | error pattern detection Excel C# | apply background color to matched cells Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

// The example loads an Excel workbook, accesses the named range "LogEntries", and iterates through its string cells. It uses a compiled regular expression (^ERROR:\s*(.+)$) to locate error messages, highlights each matching cell with a LightCoral background, prints the captured message, and saves the updated workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Ensure the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "LogEntries"
            Name logRange = workbook.Worksheets.Names["LogEntries"];
            if (logRange == null)
            {
                Console.WriteLine("Named range 'LogEntries' was not found.");
                return;
            }

            // Get the actual cell range object from the named range
            Aspose.Cells.Range range = logRange.GetRange();

            // Define the regular expression pattern to search for
            string pattern = @"^ERROR:\s*(.+)$";
            Regex regex = new Regex(pattern, RegexOptions.Compiled | RegexOptions.Multiline);

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Process only string cells
                if (cell.Type == CellValueType.IsString)
                {
                    string cellText = cell.StringValue;
                    Match match = regex.Match(cellText);
                    if (match.Success)
                    {
                        // Highlight the cell with a light red background
                        Style style = cell.GetStyle();
                        style.ForegroundColor = Color.LightCoral;
                        style.Pattern = BackgroundType.Solid;
                        cell.SetStyle(style);

                        // Output captured message
                        string capturedMessage = match.Groups[1].Value;
                        Console.WriteLine($"Match found in {cell.Name}: {capturedMessage}");
                    }
                }
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
