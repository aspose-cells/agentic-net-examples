// Title: Export rows from an Excel workbook to a CSV file using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# write each Excel row to a CSV file | C# export used range of an Excel worksheet to CSV with Aspose.Cells | How to escape commas and quotes when converting Excel to CSV using Aspose.Cells | Save large Excel data as CSV line by line in .NET with Aspose.Cells | Convert Excel workbook to CSV without loading the whole sheet into memory Aspose
// Tags: Aspose.Cells export worksheet to CSV | C# iterate Excel rows with Aspose.Cells | CSV field escaping Aspose.Cells conversion | MaxDataRow MaxDataColumn used range Aspose.Cells | StreamWriter write CSV from Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, determines the used rows and columns via MaxDataRow/MaxDataColumn, iterates each cell, applies CSV‑compatible escaping for commas, quotes, and newlines, wraps fields in double quotes when needed, and writes each row as a comma‑separated line to a CSV file using StreamWriter, while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Select the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Open a StreamWriter for the CSV output
            using (StreamWriter csvWriter = new StreamWriter(outputPath))
            {
                // Iterate over each used row
                for (int row = 0; row <= maxRow; row++)
                {
                    string[] csvFields = new string[maxCol + 1];

                    // Iterate over each used column in the current row
                    for (int col = 0; col <= maxCol; col++)
                    {
                        object cellValue = cells[row, col]?.Value;
                        string field = cellValue?.ToString() ?? string.Empty;

                        // Escape double quotes by doubling them
                        if (field.Contains("\""))
                            field = field.Replace("\"", "\"\"");

                        // Wrap the field in double quotes if it contains commas, quotes, or newlines
                        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                            field = $"\"{field}\"";

                        csvFields[col] = field;
                    }

                    // Write the CSV line
                    string csvLine = string.Join(",", csvFields);
                    csvWriter.WriteLine(csvLine);
                }
            }

            Console.WriteLine($"Conversion completed successfully. CSV saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
