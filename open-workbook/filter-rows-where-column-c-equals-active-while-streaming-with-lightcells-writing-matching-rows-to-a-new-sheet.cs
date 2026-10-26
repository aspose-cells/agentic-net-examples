// Title: C# example: filter rows with "Active" in column C and write them to a new worksheet using Aspose.Cells
// Common Searches: aspocells c# filter rows where column C equals active | copy rows with specific cell value to new sheet using Aspose.Cells in .NET | how to create a filtered workbook based on column value with Aspose.Cells C# | iterate Excel rows and export matching rows to separate file using Aspose.Cells | c# Aspose.Cells save rows that meet condition to another workbook
// Tags: filter rows by column value Aspose.Cells C# | copy matching rows to new worksheet Aspose.Cells | export filtered data to separate Excel file .NET | row iteration and conditional copy Aspose.Cells | case‑insensitive string comparison Excel column C C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, iterates through each row of the first worksheet, checks if the value in column C (zero‑based index 2) equals "Active" (case‑insensitive), copies the entire matching row to a new worksheet named "Filtered" in a separate workbook, and saves the result as output.xlsx, handling missing files and creating output directories as needed.
class FilterRowsWithLightCells
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(inputPath);
            Worksheet srcSheet = srcWorkbook.Worksheets[0];

            // Create a new workbook for the filtered result
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            destSheet.Name = "Filtered";

            int destRowIndex = 0;

            // Determine the used range in the source sheet
            int maxRow = srcSheet.Cells.MaxDataRow;
            int maxCol = srcSheet.Cells.MaxDataColumn;

            // Iterate through each row in the source sheet
            for (int row = 0; row <= maxRow; row++)
            {
                // Column C is the third column (zero‑based index = 2)
                string colCValue = srcSheet.Cells[row, 2].StringValue;

                // Check if the value equals "Active"
                if (string.Equals(colCValue, "Active", StringComparison.OrdinalIgnoreCase))
                {
                    // Copy all cells from the current source row to the destination row
                    for (int col = 0; col <= maxCol; col++)
                    {
                        var srcCell = srcSheet.Cells[row, col];
                        destSheet.Cells[destRowIndex, col].PutValue(srcCell.Value);
                    }

                    destRowIndex++;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the filtered workbook
            destWorkbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
