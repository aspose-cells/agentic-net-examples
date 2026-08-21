// Title: Convert XLS to Markdown with Aligned Tables using Aspose.Cells for .NET
// Description: Shows how to load an XLS workbook with Aspose.Cells, configure MarkdownSaveOptions (AlignColumnPadding=' ', TableHeaderType=FirstRow), and save it as a Markdown file where tables are column‑aligned and the first row is used as the header.
// Keywords: Aspose.Cells | C# XLS to Markdown | MarkdownSaveOptions | AlignColumnPadding | TableHeaderType | Excel to Markdown conversion | aligned markdown tables | export Excel as markdown | Aspose.Cells .NET example
// Common Searches: C# convert xls to markdown | Aspose.Cells markdown table alignment | export Excel as markdown with headers | MarkdownSaveOptions AlignColumnPadding example | convert legacy XLS workbook to markdown using .NET
// Developer Intent: Create a Markdown document from an XLS workbook with column‑aligned tables and header rows using Aspose.Cells.
// Use Cases: Generate documentation or static‑site pages from Excel data with readable tables. | Migrate legacy XLS reports to markdown for version‑controlled repositories. | Automate batch conversion of multiple workbooks in CI/CD pipelines. | Produce README or changelog tables directly from Excel spreadsheets.
// AI Prompts: Provide C# code that converts an XLS file to a markdown file with space‑padded column alignment and first‑row headers using Aspose.Cells. | Explain how AlignColumnPadding and TableHeaderType properties affect the generated markdown output. | Write a script to batch‑process a folder of XLS files into markdown documents with Aspose.Cells. | Describe ways to customize markdown table styling (e.g., pipe alignment) via MarkdownSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Markdown;

// Shows how to load an XLS workbook with Aspose.Cells, configure MarkdownSaveOptions (AlignColumnPadding=' ', TableHeaderType=FirstRow), and save it as a Markdown file where tables are column‑aligned and the first row is used as the header.
class Program
{
    static void Main()
    {
        // Path to the source XLS workbook
        string sourcePath = "input.xls";

        // Path where the Markdown file will be saved
        string outputPath = "output.md";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Configure Markdown save options
        // AlignColumnPadding = ' ' enables column alignment using spaces
        // TableHeaderType = FirstRow uses the first row of each sheet as the table header
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            AlignColumnPadding = ' ',
            TableHeaderType = MarkdownTableHeaderType.FirstRow
        };

        // Save the workbook as a Markdown document with aligned tables
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"Workbook converted to Markdown successfully: {outputPath}");
    }
}
