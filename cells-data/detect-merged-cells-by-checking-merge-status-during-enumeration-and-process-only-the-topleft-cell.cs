// Title: Detect merged cells in an Excel worksheet with Aspose.Cells for .NET and process only the top‑left cell of each merged range
// AI Prompts: Write C# code that loads an .xlsx workbook using Aspose.Cells, iterates over the used range, and prints the address and value of only the top‑left cell of each merged region. | Adjust an Aspose.Cells cell‑enumeration loop to skip cells that are not the first cell of a merged area while still handling regular cells. | Add robust error handling to verify the input file exists, create the output directory if missing, and save the workbook after processing cells with Aspose.Cells.
// Common Searches: C# Aspose.Cells how to check if a cell belongs to a merged range | iterate through used cells in Excel with Aspose.Cells and ignore duplicate merged cells | Aspose.Cells get top left cell of a merged region in .NET | save workbook after reading cells using Aspose.Cells C# example | detect merged cells and retrieve their values with Aspose.Cells for .NET
// Tags: Aspose.Cells merged range detection | Aspose.Cells used cells enumeration | Aspose.Cells top-left merged cell extraction | Aspose.Cells input file validation | Aspose.Cells output directory creation

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Loads an .xlsx file with Aspose.Cells, iterates through all used cells, identifies merged cells, processes only the top‑left cell of each merged range, logs cell addresses and values, validates the input file, ensures the output folder exists, and saves the workbook.
class DetectMergedCells
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists to prevent FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through all used cells
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    if (cell.IsMerged)
                    {
                        // Retrieve the merged range (use alias to avoid ambiguity with System.Range)
                        AsposeRange mergedRange = cell.GetMergedRange();

                        // Process only the top‑left cell of the merged range
                        if (mergedRange.FirstRow == row && mergedRange.FirstColumn == col)
                        {
                            Console.WriteLine($"Top‑left merged cell: {cell.Name} Value: {cell.Value}");
                        }
                    }
                    else
                    {
                        // Process regular cells
                        Console.WriteLine($"Regular cell: {cell.Name} Value: {cell.Value}");
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (optional if modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
