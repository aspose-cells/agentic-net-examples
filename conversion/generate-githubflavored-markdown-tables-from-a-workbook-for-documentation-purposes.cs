// Title: Export Excel Workbook to GitHub‑Flavored Markdown Tables with Aspose.Cells for .NET
// Description: Loads an Excel workbook, configures MarkdownSaveOptions (first‑row header, space padding, split tables by blank rows, formula evaluation) and saves it as a GitHub‑compatible Markdown file, ideal for README or documentation generation.
// Keywords: Aspose.Cells | C# | .NET | Excel to Markdown | GitHub‑flavored markdown tables | MarkdownSaveOptions | table header first row | split tables by blank rows | calculate formulas | documentation export
// Common Searches: Aspose.Cells export Excel to GitHub markdown C# | Convert Excel sheet to markdown table with Aspose.Cells | MarkdownSaveOptions split tables by blank rows example | How to evaluate formulas when saving Excel as markdown | Generate README.md tables from Excel using Aspose.Cells
// Developer Intent: Create a .md file containing GitHub‑flavored tables from an Excel workbook using Aspose.Cells.
// Use Cases: Add data tables from Excel reports directly into project README or wiki pages. | Publish calculation results with evaluated formulas to static site generators that accept markdown. | Separate multiple logical tables in a worksheet by blank rows, producing distinct markdown tables for each.
// AI Prompts: Write a C# method that uses Aspose.Cells to convert an Excel file into a GitHub‑flavored markdown file, using the first row as the header and splitting tables at blank rows. | Show how to add error handling that verifies the source Excel file exists before exporting it to markdown with Aspose.Cells. | Explain the MarkdownSaveOptions settings needed to calculate formulas, align columns with spaces, and split tables by blank rows.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Markdown;

namespace AsposeCellsDocumentation
{
    // Loads an Excel workbook, configures MarkdownSaveOptions (first‑row header, space padding, split tables by blank rows, formula evaluation) and saves it as a GitHub‑compatible Markdown file, ideal for README or documentation generation.
    public class MarkdownTableGenerator
    {
        /// <param name="excelPath">Full path to the source Excel file.</param>
        /// <param name="markdownPath">Full path where the Markdown file will be saved.</param>
        public static void GenerateMarkdown(string excelPath, string markdownPath)
        {
            try
            {
                // Verify that the source Excel file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Error: Excel file not found at '{excelPath}'.");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(excelPath);

                // Configure Markdown save options
                MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
                {
                    TableHeaderType = MarkdownTableHeaderType.FirstRow,
                    AlignColumnPadding = ' ',
                    SplitTablesByBlankRow = true,
                    CalculateFormula = true
                };

                // Save the workbook as a Markdown file
                workbook.Save(markdownPath, saveOptions);
                Console.WriteLine($"Markdown file generated at: {markdownPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while generating Markdown: {ex.Message}");
            }
        }

        // Example usage
        public static void Main()
        {
            try
            {
                string sourceExcel = @"C:\Docs\SampleData.xlsx";
                string outputMarkdown = @"C:\Docs\SampleData.md";

                GenerateMarkdown(sourceExcel, outputMarkdown);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
