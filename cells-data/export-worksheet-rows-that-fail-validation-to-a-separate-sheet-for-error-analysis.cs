// Title: C# Aspose.Cells code to export rows containing empty or whitespace cells to a separate error worksheet
// Common Searches: how to copy rows with empty cells to a new worksheet using Aspose.Cells in C# | Aspose.Cells example for exporting invalid rows to an error sheet | filter rows with missing data and save to separate sheet in .NET Excel | C# code to create an error worksheet for rows that fail validation with Aspose.Cells | export rows that fail cell validation to another sheet using Aspose.Cells library
// Tags: export rows with invalid data Aspose.Cells | copy error rows to new worksheet C# | filter empty cells in Excel workbook .NET | create error worksheet using Aspose.Cells | row validation and export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, adds a worksheet named "ErrorRows", scans each row of the first sheet for cells that are null or contain only whitespace, copies any row with such invalid cells to the error sheet, and saves the result as output.xlsx while handling missing files and runtime exceptions.
class ExportInvalidRows
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (source data)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Add a new worksheet to store rows that fail validation
            Worksheet errorSheet = workbook.Worksheets.Add("ErrorRows");

            int errorRowIndex = 0;

            // Determine the total number of rows and columns in the source sheet
            int totalRows = sourceSheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based
            int totalCols = sourceSheet.Cells.MaxDataColumn + 1;

            // Iterate through each row in the source sheet
            for (int row = 0; row < totalRows; row++)
            {
                bool hasInvalidCell = false;

                // Simple validation: treat empty or whitespace cells as invalid
                for (int col = 0; col < totalCols; col++)
                {
                    Cell cell = sourceSheet.Cells[row, col];
                    if (cell.Value == null || string.IsNullOrWhiteSpace(cell.StringValue))
                    {
                        hasInvalidCell = true;
                        break;
                    }
                }

                // If the row contains invalid data, copy it to the error sheet
                if (hasInvalidCell)
                {
                    for (int col = 0; col < totalCols; col++)
                    {
                        errorSheet.Cells[errorRowIndex, col].Value = sourceSheet.Cells[row, col].Value;
                    }
                    errorRowIndex++;
                }
            }

            // Save the workbook with the new error sheet
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
