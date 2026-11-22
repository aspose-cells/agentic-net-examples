// Title: How to export rows that fail validation to a separate worksheet for error reporting with Aspose.Cells in C#
// Common Searches: Aspose.Cells C# copy rows with empty cells to a new sheet for error log | C# generate an error report worksheet by filtering invalid rows using Aspose.Cells | Export rows containing null or whitespace values to another worksheet with Aspose.Cells | Create an error report tab in Excel from validation failures using Aspose.Cells C# | How to programmatically move rows that fail data validation to a separate sheet in C#
// Tags: Aspose.Cells filter rows by empty cells | Aspose.Cells transfer invalid rows to new sheet | Aspose.Cells generate error report worksheet | C# export validation failures to separate worksheet | Aspose.Cells row-level validation error extraction

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, scans each row for cells that are null or whitespace, collects rows with such invalid cells, creates a new worksheet named "ErrorReport", copies the header and the identified error rows into it, and saves the workbook, providing a ready‑to‑use error‑report sheet.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output_with_errors.xlsx";

        try
        {
            // Ensure the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (assumed to contain the data)
            Worksheet sourceSheet = workbook.Worksheets[0];
            Cells cells = sourceSheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;      // zero‑based index of the last used row
            int maxCol = cells.MaxDataColumn;   // zero‑based index of the last used column

            // Simple validation: treat a cell as invalid if its value is null or empty
            HashSet<int> errorRows = new HashSet<int>();
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    object val = cells[row, col].Value;
                    if (val == null || string.IsNullOrWhiteSpace(val.ToString()))
                    {
                        errorRows.Add(row);
                        break; // move to next row after first invalid cell is found
                    }
                }
            }

            // Create a new worksheet to hold the error report
            Worksheet errorSheet = workbook.Worksheets.Add("ErrorReport");

            // Copy the header row (assumed to be the first row) to the error sheet
            for (int col = 0; col <= maxCol; col++)
            {
                errorSheet.Cells[0, col].PutValue(cells[0, col].Value);
            }

            // Copy each row that has validation errors to the error sheet
            int destRow = 1; // start after header
            foreach (int srcRow in errorRows)
            {
                // Skip the header row itself if it appears in the error set
                if (srcRow == 0) continue;

                for (int col = 0; col <= maxCol; col++)
                {
                    errorSheet.Cells[destRow, col].PutValue(cells[srcRow, col].Value);
                }
                destRow++;
            }

            // Save the workbook with the new error reporting worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Processing completed. Output saved to '{outputPath}'.");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
