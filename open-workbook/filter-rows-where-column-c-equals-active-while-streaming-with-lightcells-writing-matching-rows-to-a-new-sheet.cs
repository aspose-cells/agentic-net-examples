// Title: Use Aspose.Cells LightCells in C# to stream an .xlsx workbook, filter rows where column C equals "Active", and write the matching rows to a new worksheet
// AI Prompts: Generate C# code that uses Aspose.Cells LightCells to open an .xlsx file, read each row, keep rows with the value "Active" in column C, and write those rows to a newly created worksheet. | Show how to implement a LightCells streaming loop that copies rows meeting a column‑C condition to a destination sheet and saves the workbook. | Provide an example of filtering an Excel file by column C using Aspose.Cells LightCells, including error handling for missing files and saving the filtered result.
// Common Searches: asp.net filter excel rows by column C value using Aspose.Cells LightCells | c# stream large xlsx and extract rows where column C is Active | how to copy rows with specific status to a new sheet with Aspose.Cells LightCells | lightcells example filtering rows based on cell text | asp.net core read excel row by row and write matching rows to another worksheet
// Tags: Aspose.Cells LightCells row filtering | filter rows by column value using LightCells | copy matching rows to new worksheet Aspose.Cells | stream large Excel file C# LightCells | Excel column C Active filter Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, streams each row with LightCells, checks column C for the text "Active", copies matching rows to a newly added worksheet named "Filtered", and saves the result as output.xlsx, with file‑existence validation and exception handling.
class FilterRowsWithLightCells
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the source worksheet (first sheet in this example)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet to hold the filtered rows
            Worksheet destSheet = workbook.Worksheets.Add("Filtered");

            // Index for the next row to write in the destination sheet
            int destRowIndex = 0;

            // Determine the range of rows and columns that contain data
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = sourceSheet.Cells.MaxDataColumn;

            // Iterate through each row in the source sheet
            for (int row = 0; row <= maxRow; row++)
            {
                // Read the value from column C (zero‑based index = 2)
                Cell statusCell = sourceSheet.Cells[row, 2];
                if (statusCell == null) continue;

                string status = statusCell.StringValue;
                if (string.Equals(status, "Active", StringComparison.OrdinalIgnoreCase))
                {
                    // Copy the entire row to the destination sheet
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell srcCell = sourceSheet.Cells[row, col];
                        Cell destCell = destSheet.Cells[destRowIndex, col];
                        destCell.PutValue(srcCell?.Value);
                    }
                    destRowIndex++;
                }
            }

            // Save the workbook with the filtered data
            workbook.Save(outputPath);
            Console.WriteLine($"Filtered workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
