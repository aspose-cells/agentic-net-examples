// Title: Truncate Excel cell strings longer than 50 characters with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells truncate cell text to a specific length | How to limit Excel cell string length to 50 characters using Aspose.Cells | Iterate over used range and shorten long cell values in .NET | Remove characters beyond 50 in Excel cells with Aspose.Cells library | Save workbook after trimming long text in cells using C#
// Tags: truncate cell string Aspose.Cells C# | limit Excel cell text length .NET | iterate used range cells Aspose.Cells | trim long cell values C# | save workbook after modifying cell content Aspose.Cells | maxdisplayrange cell iteration .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, walks through each worksheet's used range, checks every string cell, and replaces any value exceeding 50 characters with its first 50 characters before saving the modified file.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                if (usedRange == null)
                    continue; // Skip empty sheets

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startCol = usedRange.FirstColumn;
                int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Loop through each cell in the used range
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string values
                        if (cell.Type == CellValueType.IsString)
                        {
                            string text = cell.StringValue;
                            if (!string.IsNullOrEmpty(text) && text.Length > 50)
                            {
                                // Truncate to 50 characters
                                cell.PutValue(text.Substring(0, 50));
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
