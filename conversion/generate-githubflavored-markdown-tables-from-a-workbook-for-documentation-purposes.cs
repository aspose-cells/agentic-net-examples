// Title: Create GitHub‑Flavored Markdown Tables from an Excel Workbook Using Aspose.Cells for .NET (C#)
// Common Searches: C# code to export each worksheet of an Excel file as a GitHub markdown table with Aspose.Cells | how to generate documentation markdown from an .xlsx workbook in .NET | Aspose.Cells example for converting Excel sheets to markdown tables for GitHub | escape pipe characters when writing Excel data to markdown in C# | write Excel worksheet data to a .md file using Aspose.Cells library | convert Excel workbook to markdown documentation without losing formatting in C#
// Tags: Aspose.Cells generate markdown tables | C# export Excel to GitHub markdown | markdown table creation from workbook worksheets | escape pipe characters in markdown output | iterate Excel worksheets for documentation | write workbook data to .md file using Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMarkdownExport
{
    // The sample loads an Excel workbook with Aspose.Cells, loops through every worksheet, builds GitHub‑flavored Markdown tables (header, separator, rows), escapes markdown‑sensitive characters such as pipes, skips empty sheets, and writes the combined markdown to a .md file for documentation purposes.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel workbook
            const string inputPath = "input.xlsx";

            // Path to the generated Markdown file
            const string outputPath = "WorkbookDocumentation.md";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook (Aspose.Cells handles all supported formats)
                var workbook = new Workbook(inputPath);

                // StringBuilder will hold the complete Markdown content
                var mdBuilder = new StringBuilder();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Add a level‑2 heading for the worksheet name
                    mdBuilder.AppendLine($"## {sheet.Name}");
                    mdBuilder.AppendLine();

                    // Determine the used range of the sheet
                    var cells = sheet.Cells;
                    int maxRow = cells.MaxDataRow;      // zero‑based index of the last row with data
                    int maxCol = cells.MaxDataColumn;   // zero‑based index of the last column with data

                    if (maxRow < 0 || maxCol < 0)
                    {
                        // Empty sheet – skip further processing
                        mdBuilder.AppendLine("_(empty sheet)_");
                        mdBuilder.AppendLine();
                        continue;
                    }

                    // Build the header row (first row of the sheet)
                    var headerBuilder = new StringBuilder("|");
                    for (int c = 0; c <= maxCol; c++)
                    {
                        string headerText = cells[0, c].StringValue?.Trim() ?? string.Empty;
                        headerBuilder.Append($"{EscapeMarkdown(headerText)}|");
                    }
                    mdBuilder.AppendLine(headerBuilder.ToString());

                    // Build the separator line required by GitHub‑flavored Markdown
                    var separatorBuilder = new StringBuilder("|");
                    for (int c = 0; c <= maxCol; c++)
                    {
                        separatorBuilder.Append("---|");
                    }
                    mdBuilder.AppendLine(separatorBuilder.ToString());

                    // Append data rows (starting from row index 1)
                    for (int r = 1; r <= maxRow; r++)
                    {
                        var rowBuilder = new StringBuilder("|");
                        for (int c = 0; c <= maxCol; c++)
                        {
                            string cellText = cells[r, c].StringValue?.Trim() ?? string.Empty;
                            rowBuilder.Append($"{EscapeMarkdown(cellText)}|");
                        }
                        mdBuilder.AppendLine(rowBuilder.ToString());
                    }

                    // Add a blank line after each table for readability
                    mdBuilder.AppendLine();
                }

                // Write the generated Markdown to the output file
                File.WriteAllText(outputPath, mdBuilder.ToString(), Encoding.UTF8);

                Console.WriteLine($"Markdown documentation generated at: {Path.GetFullPath(outputPath)}");
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

            // Replace backslash first to avoid double escaping
            text = text.Replace("\\", "\\\\");
            // Escape pipe characters used as column separators
            return text.Replace("|", "\\|");
        }
    }
}
