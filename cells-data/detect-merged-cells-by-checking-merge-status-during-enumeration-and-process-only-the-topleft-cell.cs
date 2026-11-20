// Title: Detect merged cells and process only the top‑left cell of each merged range while iterating a worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate worksheet and identify merged cells | How to get only the first cell of a merged range in Aspose.Cells | Skip duplicate cells when looping through merged areas with Aspose.Cells .NET | Retrieve merged range information using IsMerged and GetMergedRange in C# | Enumerate used range cells and handle merged blocks in Aspose.Cells
// Tags: detect merged cells Aspose.Cells C# | process top-left cell of merged range Aspose.Cells | enumerate used range cells Aspose.Cells .NET | skip duplicate merged cells during iteration | IsMerged GetMergedRange API usage | handle merged areas in worksheet Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a workbook, iterates over every cell in the used range, checks each cell's IsMerged flag, obtains its merged range with GetMergedRange, and processes only the top‑left cell of each merged block before saving the workbook.
class Program
{
    static void Main()
    {
        // Define input and output file paths (replace with actual paths or arguments)
        string inputPath = "{InputFile}";
        string outputPath = "{OutputFile}";

        try
        {
            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook.");
                workbook = new Workbook();
            }

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Determine the used range limits
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Enumerate all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell belongs to a merged range
                    if (cell.IsMerged)
                    {
                        // Retrieve the merged range that contains this cell
                        Aspose.Cells.Range mergedRange = cell.GetMergedRange();

                        // Process only the top‑left cell of the merged range
                        if (mergedRange.FirstRow == row && mergedRange.FirstColumn == col)
                        {
                            // Example processing: output the cell address and its value
                            Console.WriteLine($"Top‑left merged cell: {cell.Name} = {cell.Value}");
                        }
                    }
                    else
                    {
                        // Process non‑merged cells as needed (optional)
                        // Console.WriteLine($"Regular cell: {cell.Name} = {cell.Value}");
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
