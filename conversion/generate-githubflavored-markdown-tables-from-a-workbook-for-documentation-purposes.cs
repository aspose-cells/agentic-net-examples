using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markdown;

namespace AsposeCellsDocumentation
{
    public class MarkdownTableGenerator
    {
        public static void Run()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure Markdown save options
                MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
                {
                    // Use the first row of each sheet as the table header
                    TableHeaderType = MarkdownTableHeaderType.FirstRow,

                    // Align columns with space padding for readability
                    AlignColumnPadding = ' ',

                    // Split tables by blank rows
                    SplitTablesByBlankRow = true,

                    // Calculate formulas before export
                    CalculateFormula = true,

                    // Use UTF‑8 encoding
                    Encoding = Encoding.UTF8
                };

                // Output markdown file path
                string outputPath = "DocumentationTables.md";

                // Save as markdown
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Markdown tables have been generated and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}