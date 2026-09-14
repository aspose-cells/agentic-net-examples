// Title: Batch replace old domain URLs with a new domain in the "Links" named range using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook with Aspose.Cells, retrieve the named range "Links", replace all occurrences of "oldexample.com" with "newexample.com" in each cell's text, and save the updated file. | Iterate over every cell in a specific named range using Aspose.Cells in C#, perform a substring substitution on string values, and write the modified workbook to a new location. | Create the output directory if it does not exist, then save the workbook after updating URLs in the "Links" range with Aspose.Cells.
// Common Searches: replace domain in Excel named range using Aspose.Cells C# | C# Aspose.Cells update URLs in specific range | how to change hyperlink domain in a named range with Aspose.Cells | batch update cell text in Excel workbook Aspose.Cells .NET
// Tags: Aspose.Cells modify cell strings in named range | C# update URLs in Excel range | substring replace in cell values Aspose.Cells | save workbook after named range edit C# | named range Links domain update

using Aspose.Cells;
using System;
using System.IO;
using AsposeRange = Aspose.Cells.Range;

// The example loads "input.xlsx", accesses the named range "Links", replaces every occurrence of the old domain "oldexample.com" with "newexample.com" in string cells, ensures the output folder exists, and saves the modified workbook as "output.xlsx", handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "Links"
            AsposeRange linksRange = workbook.Worksheets.GetRangeByName("Links");
            if (linksRange == null)
            {
                Console.WriteLine("Named range 'Links' not found.");
                return;
            }

            // Define the old and new domain strings
            string oldDomain = "oldexample.com";
            string newDomain = "newexample.com";

            // Iterate through each cell in the named range and replace URLs
            int startRow = linksRange.FirstRow;
            int startColumn = linksRange.FirstColumn;
            int rowCount = linksRange.RowCount;
            int columnCount = linksRange.ColumnCount;
            Worksheet ws = linksRange.Worksheet;

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    Cell cell = ws.Cells[startRow + i, startColumn + j];
                    if (cell.Value != null && cell.Value is string text && text.Contains(oldDomain))
                    {
                        string updatedText = text.Replace(oldDomain, newDomain);
                        cell.PutValue(updatedText);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
