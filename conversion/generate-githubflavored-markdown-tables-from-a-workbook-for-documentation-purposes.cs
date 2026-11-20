// Title: Generate GitHub‑Flavored Markdown Tables from an Excel Workbook using Aspose.Cells for .NET
// Common Searches: c# convert excel worksheet to markdown table with Aspose.Cells | how to export Excel data to GitHub markdown using Aspose.Cells .NET | Aspose.Cells generate markdown documentation from .xlsx file | save workbook as .md file with Aspose.Cells C# example | escape pipe characters when creating markdown tables from Excel in C#
// Tags: Aspose.Cells workbook to markdown conversion | export worksheet as GitHub markdown table .NET | markdown table generation from Excel using Aspose.Cells | escape markdown pipe characters in C# | write markdown file with UTF8 encoding C#

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMarkdownExport
{
    // This C# example loads an Excel workbook with Aspose.Cells, iterates each worksheet, and builds GitHub‑compatible markdown tables using the first row as headers. It escapes pipe and backtick characters, skips empty sheets, creates the output directory if needed, and writes the resulting markdown to a UTF‑8 .md file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file
                string excelPath = @"C:\Docs\Sample.xlsx";

                // Path to the output Markdown file
                string markdownPath = @"C:\Docs\Sample.md";

                // Verify that the source file exists
                if (!File.Exists(excelPath))
                {
                    Console.Error.WriteLine($"Error: The Excel file \"{excelPath}\" was not found.");
                    return;
                }

                // Load the workbook (lifecycle rule – use Aspose.Cells load API)
                Workbook workbook = new Workbook(excelPath);

                // Build the markdown content for all worksheets
                string markdown = GenerateMarkdownFromWorkbook(workbook);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(markdownPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the markdown text (lifecycle rule – use standard file I/O)
                File.WriteAllText(markdownPath, markdown, Encoding.UTF8);

                Console.WriteLine("Markdown documentation generated successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        /// <param name="workbook">The Aspose.Cells workbook.</param>
        /// <returns>A string containing the markdown representation.</returns>
        private static string GenerateMarkdownFromWorkbook(Workbook workbook)
        {
            StringBuilder sb = new StringBuilder();

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Worksheet title as a second‑level heading
                sb.AppendLine($"## {sheet.Name}");
                sb.AppendLine();

                // Determine the used range (non‑empty cells)
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxColumn = cells.MaxDataColumn;

                // If the sheet is empty, skip it
                if (maxRow < 0 || maxColumn < 0)
                {
                    sb.AppendLine("_Empty sheet_");
                    sb.AppendLine();
                    continue;
                }

                // Build header row – use the first row (row 0) as column titles
                sb.Append("|");
                for (int col = 0; col <= maxColumn; col++)
                {
                    string header = cells[0, col].StringValue;
                    // If header is empty, use generic column name
                    if (string.IsNullOrEmpty(header))
                        header = $"Column{col + 1}";
                    sb.Append($"{EscapeMarkdown(header)}|");
                }
                sb.AppendLine();

                // Separator row required by GitHub markdown tables
                sb.Append("|");
                for (int col = 0; col <= maxColumn; col++)
                {
                    sb.Append("---|");
                }
                sb.AppendLine();

                // Data rows (starting from row 1)
                for (int row = 1; row <= maxRow; row++)
                {
                    sb.Append("|");
                    for (int col = 0; col <= maxColumn; col++)
                    {
                        string cellText = cells[row, col].StringValue;
                        sb.Append($"{EscapeMarkdown(cellText)}|");
                    }
                    sb.AppendLine();
                }

                sb.AppendLine(); // Blank line after each table
            }

            return sb.ToString();
        }

        /// <summary>
        /// Escapes pipe characters and backticks to keep markdown tables well‑formed.
        /// </summary>
        private static string EscapeMarkdown(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            // Replace pipe characters with escaped version
            text = text.Replace("|", "\\|");
            // Replace backticks to avoid code fences inside cells
            text = text.Replace("`", "\\`");
            // Trim trailing spaces to avoid markdown rendering issues
            return text.Trim();
        }
    }
}
