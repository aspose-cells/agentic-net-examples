// Title: Export an Excel worksheet to a Markdown table with header alignment using Aspose.Cells for .NET (C#)
// Common Searches: C# how to convert an Excel worksheet to a markdown table with column alignment using Aspose.Cells | Aspose.Cells export first sheet to markdown preserving left, center, right alignment | Generate markdown table from .xlsx file in .NET while keeping header formatting | Save Excel data as markdown with pipe escaping and alignment rows in C# | Aspose.Cells markdown export example with header style detection
// Tags: Aspose.Cells export worksheet to markdown | markdown table generation from Excel in C# | preserve cell alignment in markdown using Aspose | pipe character escaping for markdown tables | C# convert .xlsx to markdown with style based alignment | Aspose.Cells style detection for markdown alignment

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsMarkdownExport
{
    // The sample loads an Excel workbook with Aspose.Cells, selects the first worksheet, determines the used range, and builds a Markdown table. It creates a header row, escapes pipe characters, adds an alignment row based on each header cell's horizontal alignment (left, center, right), writes data rows, and saves the result to a .md file, ensuring the output directory exists.
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path
            string excelPath = @"C:\Data\Sample.xlsx";

            // Output Markdown file path
            string markdownPath = @"C:\Data\Sample.md";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: The Excel file was not found at '{excelPath}'.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Choose the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Build the markdown table
                StringBuilder md = new StringBuilder();

                // ---- Header row ----
                for (int col = 0; col <= maxCol; col++)
                {
                    string header = cells[0, col].StringValue ?? string.Empty;
                    header = header.Replace("|", "\\|");
                    md.Append("| ").Append(header).Append(' ');
                }
                md.AppendLine("|");

                // Alignment row (---, :---, ---:, :---:)
                for (int col = 0; col <= maxCol; col++)
                {
                    string align = "---";
                    Style style = cells[0, col].GetStyle();

                    switch (style.HorizontalAlignment)
                    {
                        case TextAlignmentType.Left:
                            align = ":---";
                            break;
                        case TextAlignmentType.Center:
                            align = ":---:";
                            break;
                        case TextAlignmentType.Right:
                            align = "---:";
                            break;
                        default:
                            align = ":---";
                            break;
                    }

                    md.Append("| ").Append(align).Append(' ');
                }
                md.AppendLine("|");

                // ---- Data rows ----
                for (int row = 1; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        string value = cells[row, col].StringValue ?? string.Empty;
                        value = value.Replace("|", "\\|");
                        md.Append("| ").Append(value).Append(' ');
                    }
                    md.AppendLine("|");
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(markdownPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write the markdown content to file
                File.WriteAllText(markdownPath, md.ToString(), Encoding.UTF8);

                Console.WriteLine("Markdown table exported successfully to:");
                Console.WriteLine(markdownPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
