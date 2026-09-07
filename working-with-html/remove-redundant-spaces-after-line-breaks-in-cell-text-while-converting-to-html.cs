// Title: Remove extra spaces after line breaks in Excel cells before exporting to HTML using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through all string cells in an Aspose.Cells workbook, uses a regular expression to delete spaces that follow newline characters, and then saves the workbook as HTML. | Update an Aspose.Cells HTML export routine to strip whitespace occurring after '\r\n' or '\n' line breaks in cell contents.
// Common Searches: aspocells c# remove spaces after newline when saving as html | how to clean Excel cell text line break whitespace before html export using Aspose.Cells | regex to trim spaces after line breaks in Excel cells with Aspose.Cells .NET | remove trailing spaces after line breaks in workbook cells during HTML conversion Aspose.Cells
// Tags: Aspose.Cells regex newline whitespace cleanup | trim cell string values before HTML conversion | remove redundant line break spaces in Excel workbook | C# iterate worksheet cells for whitespace normalization | HTML export whitespace handling with Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHtmlConversion
{
    // The program loads an Excel workbook, walks through each worksheet's used range, applies a regular expression to replace line‑break characters followed by spaces with just the line break in string cells, updates any modified cells, and finally saves the workbook as an HTML file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
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
                    AsposeRange usedRange = sheet.Cells.MaxDisplayRange;
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

                            // Process only cells that contain string data
                            if (cell.Type == CellValueType.IsString && !string.IsNullOrEmpty(cell.StringValue))
                            {
                                // Remove redundant spaces after line breaks
                                // Pattern: line break (\r\n or \n) followed by one or more spaces
                                // Replacement: keep only the line break
                                string cleaned = Regex.Replace(cell.StringValue, @"(\r?\n)\s+", "$1");

                                // Update the cell value if changes were made
                                if (!cleaned.Equals(cell.StringValue, StringComparison.Ordinal))
                                {
                                    cell.PutValue(cleaned);
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

                // Save the workbook as HTML
                workbook.Save(outputPath, SaveFormat.Html);
                Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
