// Title: Convert XLS to Markdown with Aligned Tables using Aspose.Cells for .NET
// Description: Load an XLS workbook with Aspose.Cells, set MarkdownSaveOptions to pad columns and use the first row as a header, then save the workbook as a Markdown file containing properly aligned tables.
// Keywords: Aspose.Cells XLS to Markdown | C# markdown table alignment | MarkdownSaveOptions AlignColumnPadding | Excel to markdown conversion | TableHeaderType FirstRow
// Common Searches: export Excel to markdown with column alignment c# | Aspose.Cells AlignColumnPadding example | convert xls file to markdown table using Aspose | c# save workbook as markdown with headers
// Developer Intent: Generate a Markdown document from an XLS workbook where table columns are uniformly aligned.
// Use Cases: Create technical documentation that includes neatly formatted data tables extracted from legacy XLS files. | Automate markdown report generation for static site generators or GitHub README files. | Produce version‑controlled data snapshots by converting Excel worksheets to aligned markdown tables.
// AI Prompts: Write C# code that reads an .xls file and saves it as a markdown file with column alignment using Aspose.Cells. | Explain the impact of AlignColumnPadding and TableHeaderType on the markdown output produced by Aspose.Cells. | Provide a step‑by‑step guide to convert multiple worksheets into separate aligned markdown tables.

using System;
using Aspose.Cells;
using Aspose.Cells.Markdown;

// Load an XLS workbook with Aspose.Cells, set MarkdownSaveOptions to pad columns and use the first row as a header, then save the workbook as a Markdown file containing properly aligned tables.
class Program
{
    static void Main()
    {
        // Path to the source XLS workbook
        string sourcePath = "input.xls";

        // Desired path for the generated Markdown file
        string outputPath = "output.md";

        // Load the XLS workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure Markdown save options
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Enable column alignment by padding with spaces
            AlignColumnPadding = ' ',
            // Use the first row as the table header (optional)
            TableHeaderType = MarkdownTableHeaderType.FirstRow
        };

        // Save the workbook as a Markdown document with aligned tables
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Conversion completed successfully.");
    }
}
