// Title: Convert an XLS workbook to a Markdown file with aligned tables using Aspose.Cells in C#
// Common Searches: c# Aspose.Cells export XLS worksheet to markdown table with column alignment | how to generate markdown tables from an Excel .xls file using Aspose.Cells | convert multiple sheets from an XLS workbook to markdown headings and tables in .NET | escape pipe characters when creating markdown tables from Excel data in C#
// Tags: Aspose.Cells XLS to markdown conversion | generate markdown tables from Excel worksheets | escape markdown pipe characters in C# | left-aligned markdown table columns | iterate workbook worksheets Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace XlsToMarkdown
{
    // The example loads an XLS workbook with Aspose.Cells, iterates each worksheet, and builds a Markdown document where each sheet is introduced by a heading followed by a table. It writes the header row, a left‑aligned separator row, and all data rows, escaping backslashes and pipe characters to preserve table formatting, then saves the output to a .md file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLS file
            string inputPath = "input.xls";

            // Path to the output Markdown file
            string outputPath = "output.md";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook (Aspose.Cells handles XLS format)
                Workbook workbook = new Workbook(inputPath);

                // StringBuilder to accumulate the markdown content
                StringBuilder markdown = new StringBuilder();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add a heading for the worksheet name
                    markdown.AppendLine($"## {sheet.Name}");
                    markdown.AppendLine();

                    // Determine the used range of the sheet
                    Cells cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
                    int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

                    if (maxRow < 0 || maxCol < 0)
                    {
                        // Empty sheet – skip processing
                        continue;
                    }

                    // ----- Header Row -----
                    markdown.Append("|");
                    for (int col = 0; col <= maxCol; col++)
                    {
                        string header = cells[0, col].StringValue ?? string.Empty;
                        markdown.Append($" {EscapeMarkdown(header)} |");
                    }
                    markdown.AppendLine();

                    // ----- Alignment Row -----
                    // Using left alignment for all columns (---). Change to ":---:" for center, "---:" for right.
                    markdown.Append("|");
                    for (int col = 0; col <= maxCol; col++)
                    {
                        markdown.Append(" --- |");
                    }
                    markdown.AppendLine();

                    // ----- Data Rows -----
                    for (int row = 1; row <= maxRow; row++)
                    {
                        markdown.Append("|");
                        for (int col = 0; col <= maxCol; col++)
                        {
                            string value = cells[row, col].StringValue ?? string.Empty;
                            markdown.Append($" {EscapeMarkdown(value)} |");
                        }
                        markdown.AppendLine();
                    }

                    // Add an empty line after each table for readability
                    markdown.AppendLine();
                }

                // Write the markdown content to the output file
                File.WriteAllText(outputPath, markdown.ToString());

                Console.WriteLine($"Conversion completed. Markdown saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static string EscapeMarkdown(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            // Escape backslashes first to avoid double escaping
            text = text.Replace("\\", "\\\\");
            // Escape pipe characters which are column separators in markdown tables
            text = text.Replace("|", "\\|");
            // Trim trailing spaces to keep table alignment tidy
            return text.Trim();
        }
    }
}
