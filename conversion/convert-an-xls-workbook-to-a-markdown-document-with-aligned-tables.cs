using System;
using Aspose.Cells;
using Aspose.Cells.Markdown;

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

        // Configure Markdown save options to enable column alignment
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            AlignColumnPadding = ' ',                     // Align columns using spaces
            TableHeaderType = MarkdownTableHeaderType.FirstRow // Use the first row as table header
        };

        // Save the workbook as a Markdown document with aligned tables
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Workbook successfully converted to Markdown with aligned tables.");
    }
}