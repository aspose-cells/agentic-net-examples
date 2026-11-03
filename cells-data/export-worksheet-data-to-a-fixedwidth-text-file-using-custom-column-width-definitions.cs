// Title: Export Excel worksheet to a fixed‑width .txt file with custom column widths using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export worksheet to fixed width text file with specific column sizes | How to write Excel data to a padded fixed‑width .txt using Aspose.Cells in .NET | Aspose.Cells generate fixed‑width text output from Excel with a custom column width array | Export Excel rows to fixed‑length text columns with Aspose.Cells example in C# | Create fixed‑width text file from Excel while handling password‑protected workbook using Aspose.Cells
// Tags: Aspose.Cells fixed-width text export | custom column width array C# | pad and truncate cell values Aspose.Cells | write worksheet rows to .txt Aspose.Cells | password protected workbook handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, defines an integer array that specifies the character width for each column, iterates through all populated rows, pads or truncates each cell's string value to match the defined width, concatenates the values into a fixed‑width line, and writes the lines to an output .txt file. It also checks for the source file's existence and catches workbook‑related errors, including password‑protected files.
class ExportFixedWidth
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.txt";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (no password supplied; will throw if the file is protected)
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fixed widths (in characters) for each column to export
            int[] columnWidths = new int[] { 10, 15, 20, 12 };

            // Determine the last row with data
            int maxRow = sheet.Cells.MaxDataRow;

            // Determine the last column to process based on defined widths
            int maxCol = Math.Min(sheet.Cells.MaxDataColumn, columnWidths.Length - 1);

            // Create (or overwrite) the fixed‑width text file
            using (StreamWriter writer = new StreamWriter(outputPath))
            {
                // Iterate through each data row
                for (int row = 0; row <= maxRow; row++)
                {
                    string line = string.Empty;

                    // Process each column up to the defined widths
                    for (int col = 0; col <= maxCol; col++)
                    {
                        // Get cell text; treat null as empty string
                        string cellText = sheet.Cells[row, col].StringValue ?? string.Empty;

                        // Truncate or pad the text to match the column width
                        if (cellText.Length > columnWidths[col])
                            cellText = cellText.Substring(0, columnWidths[col]);
                        else
                            cellText = cellText.PadRight(columnWidths[col]);

                        line += cellText;
                    }

                    writer.WriteLine(line);
                }
            }

            Console.WriteLine($"Export completed successfully to \"{outputPath}\".");
        }
        catch (CellsException cex)
        {
            // Handle workbook-related errors (e.g., password protection)
            if (cex.Message != null && cex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected. Please provide a valid password.");
            }
            else
            {
                Console.WriteLine($"Aspose.Cells error: {cex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
