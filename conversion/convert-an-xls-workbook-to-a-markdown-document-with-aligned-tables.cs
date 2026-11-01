// Title: Convert an XLS workbook to a Markdown document with aligned tables using Aspose.Cells for .NET
// Common Searches: C# convert .xls file to markdown table with column alignment using Aspose.Cells | How to export Excel worksheets as aligned markdown tables in .NET | Generate markdown from a workbook with padded cells using Aspose.Cells | Create a markdown file from multiple Excel sheets preserving table formatting in C# | Programmatically convert legacy XLS workbook to markdown document
// Tags: Aspose.Cells XLS to Markdown conversion | aligned markdown tables from Excel data | C# compute column widths for markdown tables | export workbook worksheets as markdown | markdown table generation with Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace XlsToMarkdown
{
    // The example loads an input.xls workbook with Aspose.Cells, iterates each worksheet, calculates the maximum width of each column, and builds a Markdown document where each sheet is represented by a heading followed by an aligned table. Cells are padded based on column widths, a separator row is added after the header, and the resulting markdown is saved to output.md with basic error handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.md";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the XLS workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }

            var markdown = new StringBuilder();

            try
            {
                // Process each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add a heading for the worksheet
                    markdown.AppendLine($"## {sheet.Name}");
                    markdown.AppendLine();

                    // Determine the used range of the sheet
                    int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
                    int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last column with data

                    // Skip empty sheets
                    if (maxRow < 0 || maxCol < 0)
                    {
                        continue;
                    }

                    // Compute the maximum width of each column (for alignment)
                    int[] colWidths = new int[maxCol + 1];
                    for (int r = 0; r <= maxRow; r++)
                    {
                        for (int c = 0; c <= maxCol; c++)
                        {
                            string cellText = sheet.Cells[r, c].StringValue ?? string.Empty;
                            colWidths[c] = Math.Max(colWidths[c], cellText.Length);
                        }
                    }

                    // Local function to write a row to the markdown builder
                    void AppendRow(int rowIndex, bool isHeader = false)
                    {
                        markdown.Append("|");
                        for (int c = 0; c <= maxCol; c++)
                        {
                            string cellText = sheet.Cells[rowIndex, c].StringValue ?? string.Empty;
                            string padded = cellText.PadRight(colWidths[c]);
                            markdown.Append($" {padded} |");
                        }
                        markdown.AppendLine();

                        // After the header row, add the separator line required by Markdown tables
                        if (isHeader)
                        {
                            markdown.Append("|");
                            for (int c = 0; c <= maxCol; c++)
                            {
                                string dashes = new string('-', colWidths[c]);
                                markdown.Append($" {dashes} |");
                            }
                            markdown.AppendLine();
                        }
                    }

                    // Write the header row (first row of the sheet)
                    AppendRow(0, isHeader: true);

                    // Write the remaining data rows
                    for (int r = 1; r <= maxRow; r++)
                    {
                        AppendRow(r);
                    }

                    // Add an empty line after each table for readability
                    markdown.AppendLine();
                }

                // Save the generated Markdown to a file
                File.WriteAllText(outputPath, markdown.ToString());
                Console.WriteLine($"Conversion completed. Markdown saved to \"{outputPath}\"");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
