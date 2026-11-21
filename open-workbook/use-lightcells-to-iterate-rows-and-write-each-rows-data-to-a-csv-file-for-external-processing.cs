// Title: Export each row of an Excel worksheet to a CSV file with proper escaping using Aspose.Cells for .NET (C#)
// Common Searches: how to write Excel worksheet rows to a CSV file with Aspose.Cells in C# | Aspose.Cells export used range to CSV with custom field escaping | C# iterate over worksheet cells and generate CSV using Aspose.Cells | save Excel data as CSV handling commas and quotes with Aspose.Cells .NET | Aspose.Cells stream rows to CSV without loading the whole workbook
// Tags: Aspose.Cells export worksheet to CSV | C# iterate worksheet rows Aspose.Cells | CSV field escaping Aspose.Cells | used range extraction Aspose.Cells .NET | streaming Excel rows to CSV C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, determines the used row and column range, and iterates each cell row‑by‑row. For every cell it retrieves the string value, applies CSV escaping rules (doubling quotes and wrapping fields that contain commas, newlines, or quotes), and writes the assembled line to a CSV file via StreamWriter. It also ensures the output directory exists and handles missing input files or runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range
            int maxRow = worksheet.Cells.MaxDataRow;
            int maxCol = worksheet.Cells.MaxDataColumn;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Open a StreamWriter for the CSV output
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Iterate through each row in the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    List<string> csvFields = new List<string>();

                    // Loop through all columns in the current row
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Retrieve the cell value (can be null)
                        Cell cell = worksheet.Cells[row, col];
                        string text = cell?.StringValue ?? string.Empty;

                        // Escape double quotes by doubling them
                        if (text.Contains("\""))
                            text = text.Replace("\"", "\"\"");

                        // If the field contains commas, newlines or quotes, wrap it in quotes
                        if (text.Contains(",") || text.Contains("\n") || text.Contains("\r") || text.Contains("\""))
                            text = $"\"{text}\"";

                        csvFields.Add(text);
                    }

                    // Write the CSV line
                    writer.WriteLine(string.Join(",", csvFields));
                }
            }

            Console.WriteLine("CSV export completed.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
