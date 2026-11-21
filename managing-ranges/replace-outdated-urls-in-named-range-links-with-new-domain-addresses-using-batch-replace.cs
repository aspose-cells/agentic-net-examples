// Title: How to batch replace outdated URLs in the named range "Links" of an Excel workbook using Aspose.Cells for .NET
// Common Searches: C# code to replace old domain URLs in a specific named range with Aspose.Cells | batch update hyperlink URLs in an Excel named range using Aspose.Cells .NET | replace string values in a named range programmatically with Aspose.Cells | iterate over cells in a named range and modify URLs with Aspose.Cells C# | update Excel workbook URLs from http to https in a named range via C#
// Tags: named range batch string replace Aspose.Cells | domain URL update in Excel cells C# | Aspose.Cells iterate named range cells | modify hyperlink domain in workbook .NET | Excel workbook save after URL change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the single‑area named range "Links", iterates through each cell in that range, replaces occurrences of a specified old domain with a new domain in string cells, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the named range "Links"
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range object (single‑area named range)
            Aspose.Cells.Range linksRange = workbook.Worksheets.GetRangeByName("Links");
            if (linksRange == null)
            {
                Console.WriteLine("Named range 'Links' not found.");
                return;
            }

            // Define the old and new domain strings for replacement
            string oldDomain = "http://olddomain.com";
            string newDomain = "https://newdomain.com";

            // Determine the bounds of the range
            int firstRow = linksRange.FirstRow;
            int firstCol = linksRange.FirstColumn;
            int lastRow = firstRow + linksRange.RowCount - 1;
            int lastCol = firstCol + linksRange.ColumnCount - 1;

            Worksheet sheet = linksRange.Worksheet;

            // Iterate through all cells within the named range
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Only process string cells that contain the old domain
                    if (cell.Type == CellValueType.IsString && cell.StringValue.Contains(oldDomain))
                    {
                        string updatedValue = cell.StringValue.Replace(oldDomain, newDomain);
                        cell.PutValue(updatedValue);
                    }
                }
            }

            // Save the workbook with the updated URLs
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
