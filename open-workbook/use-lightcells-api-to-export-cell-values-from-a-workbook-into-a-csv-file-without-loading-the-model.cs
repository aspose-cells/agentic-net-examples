// Title: Export Excel worksheet cell values to a CSV file using Aspose.Cells in C# with low memory overhead
// Common Searches: how to write Excel worksheet data to CSV using Aspose.Cells C# without loading entire sheet into memory | Aspose.Cells C# export first worksheet to CSV with custom cell value handling | C# stream Excel cell values to CSV file using Aspose.Cells and avoid full workbook load | generate CSV from XLSX in C# with Aspose.Cells while handling commas and quotes | Aspose.Cells example for exporting used range of worksheet to CSV in .NET
// Tags: Aspose.Cells CSV export C# | export worksheet cells to CSV Aspose.Cells | low memory Excel to CSV conversion .NET | custom cell value escaping CSV Aspose.Cells | used range extraction CSV Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an XLSX workbook with Aspose.Cells, determines the used range of the first worksheet, and writes each cell's value to a CSV file using a StreamWriter. It creates the output directory if needed, handles empty sheets, and correctly escapes commas, newlines, and double quotes to produce a standards‑compliant CSV.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path to the destination CSV file
        string outputPath = "output.csv";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (default options load all data)
            Workbook workbook = new Workbook(inputPath);

            // Work with the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine the used range of the worksheet
            int maxRow = worksheet.Cells.MaxDataRow;          // zero‑based index of the last row with data
            int maxColumn = worksheet.Cells.MaxDataColumn;    // zero‑based index of the last column with data

            // If the sheet is empty, create an empty file and exit
            if (maxRow < 0 || maxColumn < 0)
            {
                File.WriteAllText(outputPath, string.Empty);
                Console.WriteLine($"CSV file created (empty) at: {outputPath}");
                return;
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Export the data to CSV
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        // Retrieve the cell value
                        object cellValue = worksheet.Cells[row, col].Value;
                        string text = cellValue != null ? cellValue.ToString() : string.Empty;

                        // Escape double quotes
                        if (text.Contains("\""))
                            text = text.Replace("\"", "\"\"");

                        // Enclose in quotes if the field contains commas, new lines or quotes
                        if (text.Contains(",") || text.Contains("\n") || text.Contains("\r") || text.Contains("\""))
                            text = $"\"{text}\"";

                        writer.Write(text);

                        // Write column separator except after the last column
                        if (col < maxColumn)
                            writer.Write(",");
                    }
                    // End of row
                    writer.WriteLine();
                }
            }

            Console.WriteLine($"CSV file created successfully at: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
