// Title: Replace outdated domain URLs in a named range using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells replace old domain in named range cells | How to batch update URLs in an Excel named range with Aspose.Cells | Replace http://oldexample.com with https://newexample.com in a workbook using Aspose.Cells | Iterate through named range cells and modify string values with Aspose.Cells C# | Update hyperlink text in an Excel range programmatically using Aspose.Cells
// Tags: batch replace URLs Aspose.Cells C# | named range string manipulation Aspose.Cells | update domain links Excel workbook .NET | cell value string replace Aspose.Cells | save modified workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the named range "Links", iterates over each cell, replaces occurrences of the old domain (http://oldexample.com) with the new domain (https://newexample.com) in string cells, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the input file.
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range called "Links".
            // GetRangeByName returns a Range object that represents the cells in the named range.
            Aspose.Cells.Range linksRange = workbook.Worksheets.GetRangeByName("Links");
            if (linksRange == null)
            {
                Console.WriteLine("Named range 'Links' was not found in the workbook.");
                return;
            }

            // Define the old domain to be replaced and the new domain.
            string oldDomain = "http://oldexample.com";
            string newDomain = "https://newexample.com";

            // Iterate through every cell inside the named range.
            for (int i = 0; i < linksRange.RowCount; i++)
            {
                for (int j = 0; j < linksRange.ColumnCount; j++)
                {
                    // Access the cell at the current position.
                    Cell cell = linksRange[i, j];

                    // Process only string cells that may contain URLs.
                    if (cell.Type == CellValueType.IsString)
                    {
                        string cellText = cell.StringValue;

                        // Perform batch replace of the old domain with the new one.
                        if (cellText.Contains(oldDomain))
                        {
                            string updatedText = cellText.Replace(oldDomain, newDomain);
                            cell.PutValue(updatedText);
                        }
                    }
                }
            }

            // Save the modified workbook to a new file.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
