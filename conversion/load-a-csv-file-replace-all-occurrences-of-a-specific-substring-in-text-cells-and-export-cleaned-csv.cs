// Title: Replace a specific substring in CSV text cells with Aspose.Cells for .NET and save the cleaned file
// Common Searches: Aspose.Cells C# replace text in CSV file and save as new CSV | How to load a CSV workbook with Aspose.Cells and update string values | C# example for replacing all occurrences of a word in CSV cells using Aspose.Cells | Iterate over used range of a CSV worksheet and modify cell text with Aspose.Cells .NET | Clean CSV data by substituting substrings in cells via Aspose.Cells API
// Tags: csv cell string replacement using Aspose.Cells | load csv workbook with Aspose.Cells LoadOptions | iterate used range cells Aspose.Cells | save workbook to csv format Aspose.Cells | update worksheet cell string values Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample loads a CSV file into an Aspose.Cells Workbook, walks through the used range of each worksheet, replaces every occurrence of a specified old substring with a new one in string‑type cells, ensures the output directory exists, and saves the modified workbook back to CSV.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string inputPath = "input.csv";

        // Path to the cleaned CSV file to be saved
        string outputPath = "output.csv";

        // Substring to find and its replacement
        string oldValue = "OldSubstring";
        string newValue = "NewSubstring";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the CSV file into a Workbook using CSV load options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through all worksheets (CSV typically has one sheet)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the sheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                // Determine start and end indices
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

                        // Process only text (string) cells
                        if (cell.Type == CellValueType.IsString && cell.StringValue != null)
                        {
                            // Replace all occurrences of the target substring
                            string updatedText = cell.StringValue.Replace(oldValue, newValue);
                            cell.PutValue(updatedText);
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the cleaned workbook back to CSV format
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"File saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
