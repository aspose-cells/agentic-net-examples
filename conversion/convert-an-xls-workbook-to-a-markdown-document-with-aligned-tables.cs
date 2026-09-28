// Title: How to convert an XLS workbook to a Markdown document with column-aligned tables using Aspose.Cells for .NET
// AI Prompts: Write a C# snippet that opens an .xls file with Aspose.Cells, configures the Markdown export to use space characters for padding, and writes the result to a .md file. | Demonstrate exporting an Excel worksheet to Markdown where the top row becomes the table header and each column is spaced evenly. | Build a console program that reads an existing XLS workbook and generates a Markdown file with neatly padded tables using Aspose.Cells settings.
// Common Searches: aspnet convert xls to markdown with aligned columns using Aspose.Cells | c# export excel sheet to markdown table preserving header row | how to set MarkdownSaveOptions AlignColumnPadding in Aspose.Cells | save workbook as markdown with space padding for columns c# | markdown table alignment from Excel using Aspose.Cells example
// Tags: Aspose.Cells MarkdownSaveOptions column spacing | XLS workbook to markdown conversion Aspose.Cells | C# create markdown tables from Excel data | first worksheet row as markdown table header | aligned markdown tables via Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Markdown;

// The program loads an XLS workbook with Aspose.Cells, configures MarkdownSaveOptions to align columns using spaces and to treat the first row as the table header, then saves the workbook as a Markdown file with properly aligned tables.
class Program
{
    static void Main()
    {
        // Source XLS workbook path
        string sourcePath = "input.xls";

        // Destination Markdown file path
        string outputPath = "output.md";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Set up Markdown save options with column alignment enabled
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            AlignColumnPadding = ' ',                     // Align columns using spaces
            TableHeaderType = MarkdownTableHeaderType.FirstRow // Use the first row as table header
        };

        // Save the workbook as a Markdown document
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Workbook successfully converted to Markdown with aligned tables.");
    }
}
