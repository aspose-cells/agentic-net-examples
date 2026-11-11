// Title: C# Aspose.Cells example: filter rows where column C is "Active" and copy them to a new worksheet
// Common Searches: Aspose.Cells C# filter rows where column C equals Active | How to copy rows with a specific cell value to another sheet using Aspose.Cells | C# iterate Excel rows with Aspose.Cells and write matching rows to a new worksheet | Extract rows with status Active from an Excel file using Aspose.Cells in .NET | Save filtered rows to a separate sheet with Aspose.Cells C#
// Tags: filter rows by column value Aspose.Cells | copy filtered rows to new worksheet Aspose.Cells | Aspose.Cells read Excel rows C# | Aspose.Cells write filtered data to another sheet | Aspose.Cells iterate rows with condition

using System;
using System.IO;
using Aspose.Cells;

// The code loads an Excel workbook, scans each row of the first worksheet, checks if the cell in column C contains the text "Active" (case‑insensitive), copies the entire matching row to a newly added worksheet named "Filtered", and saves the result as a new file.
class FilterRowsWithLightCells
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet (first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Destination worksheet for filtered rows
            Worksheet destSheet = workbook.Worksheets.Add("Filtered");

            int destRowIndex = 0; // Tracks the current row in the destination sheet

            // Determine the range of used rows and columns
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = sourceSheet.Cells.MaxDataColumn;

            // Iterate through each row in the source sheet
            for (int row = 0; row <= maxRow; row++)
            {
                // Column C is the third column (zero‑based index 2)
                string status = sourceSheet.Cells[row, 2].StringValue;

                // If the value in column C equals "Active", copy the entire row
                if (string.Equals(status, "Active", StringComparison.OrdinalIgnoreCase))
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        object cellValue = sourceSheet.Cells[row, col].Value;
                        destSheet.Cells[destRowIndex, col].PutValue(cellValue);
                    }

                    destRowIndex++;
                }
            }

            // Save the workbook with the filtered data to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Filtered data saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
