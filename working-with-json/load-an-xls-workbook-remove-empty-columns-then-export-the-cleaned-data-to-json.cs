// Title: Clean an XLS workbook by deleting fully empty columns and save the result as JSON with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xls file using Aspose.Cells, scans each worksheet, removes columns that contain only null cells, and writes the workbook to a JSON file. | Show how to safely verify the input file, iterate backwards through columns to avoid index shifts, and call Workbook.Save with SaveFormat.Json after the column removal process. | Provide a robust example that includes try‑catch exception handling and logs a confirmation message when the JSON export finishes successfully.
// Common Searches: C# Aspose.Cells remove fully blank columns prior to JSON export | How to remove blank columns from every sheet in an XLS file using Aspose.Cells .NET | Save cleaned Excel workbook as JSON with Aspose.Cells after column pruning | Aspose.Cells .NET example for removing columns with no data and converting XLS to JSON
// Tags: Aspose.Cells delete empty columns C# | XLS to JSON conversion after column cleanup | remove blank columns workbook Aspose.Cells | Aspose.Cells column pruning before JSON export | C# Aspose.Cells save workbook as JSON

using System;
using System.IO;
using Aspose.Cells;

// // Loads an XLS workbook, removes any columns that consist solely of null values from each worksheet, and saves the cleaned data as a JSON file using Aspose.Cells.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xls";
        const string outputPath = "output.json";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the maximum column index that contains data
                int maxColumn = cells.MaxColumn;

                // Loop backwards to avoid index shift when deleting columns
                for (int col = maxColumn; col >= 0; col--)
                {
                    bool isEmptyColumn = true;

                    // Check each cell in the current column
                    for (int row = 0; row <= cells.MaxRow; row++)
                    {
                        Cell cell = cells[row, col];

                        // Determine if the cell contains any data
                        if (cell.Type != CellValueType.IsNull)
                        {
                            isEmptyColumn = false;
                            break;
                        }
                    }

                    // Delete the column if it is completely empty
                    if (isEmptyColumn)
                    {
                        cells.DeleteColumn(col);
                    }
                }
            }

            // Export the cleaned workbook data to JSON format
            workbook.Save(outputPath, SaveFormat.Json);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
