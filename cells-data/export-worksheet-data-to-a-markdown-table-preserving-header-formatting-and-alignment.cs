// Title: Export an Excel worksheet to a Markdown table with bold headers and aligned columns using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads the first worksheet of an Excel file with Aspose.Cells, detects bold header cells, and generates a Markdown table where those headers are wrapped with ** markers. | Create a method that scans all cells to compute the maximum width per column, pads each value for left‑aligned columns, inserts the Markdown alignment separator row, and saves the output as a UTF‑8 .md file.
// Common Searches: how to convert an Excel sheet to a markdown table with Aspose.Cells in C# | preserve Excel bold header formatting when exporting to markdown using Aspose.Cells | generate aligned markdown columns from worksheet data in .NET | Aspose.Cells export worksheet to markdown file with column width calculation | C# write worksheet data to markdown with left-aligned cells
// Tags: Aspose.Cells export worksheet to markdown | calculate markdown column widths C# | preserve bold header Aspose.Cells | write UTF-8 markdown file .NET | aligned markdown table from Excel data

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMarkdownExport
{
    // The example loads a workbook, identifies the used range, computes the widest content per column, builds a Markdown table with left‑aligned cells, adds ** markers for bold header cells, inserts the alignment separator row, and writes the result to a UTF‑8 .md file while handling missing source files and creating the output directory.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string excelPath = @"C:\Data\SourceWorkbook.xlsx";

            // Path to the output Markdown file
            string markdownPath = @"C:\Data\WorksheetExport.md";

            try
            {
                // Verify that the source workbook exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Source Excel file not found: {excelPath}");
                    return;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(markdownPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Get the first worksheet (adjust index or name as needed)
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Determine the used range
                int maxRow = cells.MaxDataRow;        // zero‑based index of the last row with data
                int maxCol = cells.MaxDataColumn;     // zero‑based index of the last column with data

                // Compute the maximum width for each column (including header)
                int[] colWidths = new int[maxCol + 1];
                for (int col = 0; col <= maxCol; col++)
                {
                    int maxWidth = 0;
                    for (int row = 0; row <= maxRow; row++)
                    {
                        Cell cell = cells[row, col];
                        string text = cell.Value?.ToString() ?? string.Empty;

                        // If the header cell is bold, account for markdown ** markers
                        if (row == 0 && cell.GetStyle().Font.IsBold)
                        {
                            text = $"**{text}**";
                        }

                        if (text.Length > maxWidth)
                            maxWidth = text.Length;
                    }
                    colWidths[col] = maxWidth;
                }

                // Build the markdown table
                StringBuilder md = new StringBuilder();

                for (int row = 0; row <= maxRow; row++)
                {
                    md.Append("|");
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        string text = cell.Value?.ToString() ?? string.Empty;

                        // Apply markdown bold formatting if the header cell is bold
                        if (row == 0 && cell.GetStyle().Font.IsBold)
                        {
                            text = $"**{text}**";
                        }

                        // Pad the cell content to the column width (left‑aligned)
                        md.Append(' ');
                        md.Append(text.PadRight(colWidths[col]));
                        md.Append(' ');
                        md.Append("|");
                    }
                    md.AppendLine();

                    // After writing the header row, insert the alignment separator line
                    if (row == 0)
                    {
                        md.Append("|");
                        for (int col = 0; col <= maxCol; col++)
                        {
                            string dashes = new string('-', colWidths[col]);
                            md.Append(' ');
                            md.Append(dashes);
                            md.Append(' ');
                            md.Append("|");
                        }
                        md.AppendLine();
                    }
                }

                // Write the markdown content to file
                File.WriteAllText(markdownPath, md.ToString(), Encoding.UTF8);

                Console.WriteLine($"Worksheet exported to markdown at: {markdownPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
