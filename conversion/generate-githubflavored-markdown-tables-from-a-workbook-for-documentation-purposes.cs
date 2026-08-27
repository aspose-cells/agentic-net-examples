// Title: Generate separate GitHub‑flavored Markdown files for each worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that loads an .xlsx file, iterates through all worksheets, and saves each worksheet as an individual .md file containing a GitHub‑compatible markdown table, applying Aspose.Cells MarkdownSaveOptions with first‑row headers and formula evaluation. | Create a C# routine that builds a temporary workbook for each sheet, configures MarkdownSaveOptions (space padding, no splitting by blank rows), and exports the sheet to a markdown file named after the worksheet.
// Common Searches: C# Aspose.Cells export each Excel sheet to separate markdown file | How to save Excel worksheets as GitHub markdown tables with formula calculation in .NET | Aspose.Cells MarkdownSaveOptions settings for per‑worksheet export | Generate .md files from multiple worksheets using Aspose.Cells | Convert Excel workbook to markdown tables programmatically in C#
// Tags: Aspose.Cells per‑worksheet markdown export | C# MarkdownSaveOptions first‑row header | export Excel sheet to GitHub markdown | temporary workbook sheet copy Aspose.Cells | formula evaluation in markdown conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markdown;

namespace AsposeCellsExamples
{
    // The example loads 'input.xlsx', verifies its presence, iterates over every worksheet, creates a temporary workbook containing only the current sheet, and saves it as a markdown file (e.g., Sheet1.md) using MarkdownSaveOptions configured for first‑row headers, space padding, no table splitting, and formula calculation.
    public class GenerateMarkdownTables
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The file '{inputPath}' was not found.");

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Common markdown export options
            MarkdownSaveOptions markdownOptions = new MarkdownSaveOptions
            {
                TableHeaderType = MarkdownTableHeaderType.FirstRow,
                AlignColumnPadding = ' ',
                SplitTablesByBlankRow = false,
                CalculateFormula = true
            };

            // Export each worksheet individually
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                try
                {
                    Worksheet sheet = sourceWorkbook.Worksheets[i];

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWorkbook = new Workbook();
                    tempWorkbook.Worksheets.Clear();

                    // Add a copy of the current worksheet by name (required overload)
                    tempWorkbook.Worksheets.AddCopy(sheet.Name);

                    string mdFileName = $"{sheet.Name}.md";

                    // Save the temporary workbook as markdown
                    tempWorkbook.Save(mdFileName, markdownOptions);

                    Console.WriteLine($"Worksheet '{sheet.Name}' exported to '{mdFileName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export worksheet index {i}: {ex.Message}");
                }
            }
        }
    }
}
