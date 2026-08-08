// Title: Export an Excel worksheet to a Markdown table with header row and column alignment using Aspose.Cells for .NET
// Description: Creates a workbook, adds a header row and sample data, then saves the sheet as a .md file. The MarkdownSaveOptions treat the first row as the table header, pad columns with spaces for alignment, output displayed strings, use UTF‑8 encoding, and write Unix line endings.
// Keywords: Aspose.Cells markdown export | Excel to Markdown table .NET | MarkdownSaveOptions header row | AlignColumnPadding Aspose | CellValueFormatStrategy DisplayString | UTF-8 markdown file
// Common Searches: how to export Excel to markdown with Aspose.Cells | Aspose.Cells markdown table header alignment | save worksheet as markdown file .NET | markdown column padding Aspose.Cells
// Developer Intent: Generate a Markdown file from a worksheet where the first row becomes the table header and each column is padded for visual alignment.
// Use Cases: Generate README tables from product lists stored in Excel. | Automate markdown reports from financial or inventory spreadsheets. | Create documentation tables for GitHub, GitLab, or Confluence directly from Excel data.
// AI Prompts: Show how to right‑align numeric columns in the Markdown output with Aspose.Cells. | Demonstrate exporting multiple worksheets to separate Markdown files, each with its own header. | Explain how to change the line separator to Windows CRLF and switch to ISO‑8859‑1 encoding.

using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markdown;

// Creates a workbook, adds a header row and sample data, then saves the sheet as a .md file. The MarkdownSaveOptions treat the first row as the table header, pad columns with spaces for alignment, output displayed strings, use UTF‑8 encoding, and write Unix line endings.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row (will be used as markdown table header)
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["C1"].PutValue("Quantity");

        // Add sample data rows
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(1.99);
        sheet.Cells["C2"].PutValue(10);

        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(0.99);
        sheet.Cells["C3"].PutValue(20);

        // Configure markdown save options
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Use the first row of the worksheet as the markdown table header
            TableHeaderType = MarkdownTableHeaderType.FirstRow,
            // Enable column alignment by padding with spaces
            AlignColumnPadding = ' ',
            // Export cell values as displayed strings
            FormatStrategy = CellValueFormatStrategy.DisplayString,
            // Set encoding and line separator for the output file
            Encoding = Encoding.UTF8,
            LineSeparator = "\n"
        };

        // Save the worksheet as a markdown file with the specified options
        workbook.Save("WorksheetExport.md", saveOptions);
    }
}
