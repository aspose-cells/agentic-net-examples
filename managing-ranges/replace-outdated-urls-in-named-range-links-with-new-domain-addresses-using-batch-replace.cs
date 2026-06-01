using System;
using System.IO;
using Aspose.Cells;

namespace ReplaceUrlsInNamedRange
{
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range "Links"
                Aspose.Cells.Range linksRange = workbook.Worksheets.GetRangeByName("Links");
                if (linksRange == null)
                {
                    Console.WriteLine("Named range 'Links' not found in the workbook.");
                    return;
                }

                // Domains to replace
                string oldDomain = "oldexample.com";
                string newDomain = "newexample.com";

                // Iterate through each cell in the named range
                for (int row = 0; row < linksRange.RowCount; row++)
                {
                    for (int col = 0; col < linksRange.ColumnCount; col++)
                    {
                        // Access the cell relative to the range's first cell
                        Cell cell = linksRange[row, col];

                        // Process only string cells
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;
                            if (original.Contains(oldDomain))
                            {
                                string updated = original.Replace(oldDomain, newDomain);
                                cell.PutValue(updated);
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}