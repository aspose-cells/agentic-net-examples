// Title: Replace a specific substring in CSV text cells using Aspose.Cells for .NET and export the cleaned file
// Common Searches: Aspose.Cells replace text in CSV cells C# | C# load CSV with Aspose.Cells and substitute a string in all rows | How to clean CSV data by replacing a substring using Aspose.Cells .NET | Saving modified CSV after string replacement with Aspose.Cells in C# | Iterate over used range of a CSV worksheet with Aspose.Cells to update cell values
// Tags: Aspose.Cells CSV substring replacement | C# iterate worksheet used range Aspose.Cells | replace string in cell values Aspose.Cells | export workbook to CSV format Aspose.Cells | ensure output directory before saving CSV C#

using System;
using System.IO;
using Aspose.Cells;

// The program loads a CSV file into an Aspose.Cells Workbook, iterates over the used range of the first worksheet, replaces occurrences of a specified old substring with a new one in string‑type cells, creates the output directory if needed, and saves the modified workbook back to CSV.
class CsvCleaner
{
    static void Main()
    {
        // Paths for input and output CSV files
        string inputPath = "input.csv";
        string outputPath = "output.csv";

        // Substring to find and its replacement
        string oldValue = "OldSubstring";
        string newValue = "NewSubstring";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the CSV file into a Workbook (Aspose.Cells auto-detects CSV format)
            Workbook workbook = new Workbook(inputPath);

            // CSV files are loaded into a single worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range of cells
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            int startRow = usedRange.FirstRow;
            int startColumn = usedRange.FirstColumn;
            int endRow = startRow + usedRange.RowCount - 1;
            int endColumn = startColumn + usedRange.ColumnCount - 1;

            // Iterate through all used cells
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = sheet.Cells[row, col];

                    // Process only text cells
                    if (cell.Type == CellValueType.IsString && cell.Value != null)
                    {
                        string text = cell.StringValue;
                        if (text.Contains(oldValue))
                        {
                            // Replace the specified substring
                            cell.PutValue(text.Replace(oldValue, newValue));
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

            // Save the modified workbook back to CSV format
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Cleaned CSV saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
