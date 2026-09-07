// Title: Remove empty columns from an XLS workbook and export the cleaned data to JSON using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xls file with Aspose.Cells, scans each worksheet for columns that contain only blank or whitespace cells, deletes those columns, and then exports the workbook to a JSON file. | Adjust the column‑deletion routine to ignore hidden columns while still using JsonSaveOptions to write the cleaned workbook to JSON. | Add logging that records the indexes of removed columns for each worksheet before saving the workbook as JSON with Aspose.Cells.
// Common Searches: Aspose.Cells C# remove columns that are completely empty from an XLS file | How to export a cleaned Excel sheet to JSON using Aspose.Cells .NET | C# code to delete blank columns in each worksheet before saving as JSON with Aspose.Cells | Using JsonSaveOptions to convert an XLS workbook to JSON after column cleanup | Detect and delete empty columns in Aspose.Cells workbook programmatically
// Tags: delete empty columns Aspose.Cells | json export Aspose.Cells | xls workbook processing C# | column cleanup before JSON conversion | JsonSaveOptions usage

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// // Loads an .xls workbook, removes any columns that consist solely of empty or whitespace cells from each worksheet, and saves the cleaned data to a JSON file using Aspose.Cells JsonSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.json";

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

                // Determine the used range dimensions
                int maxColumn = cells.MaxColumn + 1; // total number of columns used
                int maxRow = cells.MaxRow + 1;       // total number of rows used

                List<int> emptyColumns = new List<int>();

                // Identify columns where every cell is empty or whitespace
                for (int col = 0; col < maxColumn; col++)
                {
                    bool isEmpty = true;
                    for (int row = 0; row < maxRow; row++)
                    {
                        object value = cells[row, col].Value;
                        if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                        {
                            isEmpty = false;
                            break;
                        }
                    }
                    if (isEmpty)
                    {
                        emptyColumns.Add(col);
                    }
                }

                // Delete empty columns starting from the rightmost to keep indices correct
                for (int i = emptyColumns.Count - 1; i >= 0; i--)
                {
                    cells.DeleteColumn(emptyColumns[i]);
                }
            }

            // Export the cleaned workbook to JSON format using JsonSaveOptions
            JsonSaveOptions jsonOptions = new JsonSaveOptions();
            workbook.Save(outputPath, jsonOptions);
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
