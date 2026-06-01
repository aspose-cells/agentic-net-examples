using System;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Markdown;

class Program
{
    static void Main()
    {
        // Path to the source XLS workbook
        string sourcePath = "input.xls";

        // Desired path for the generated Markdown file
        string outputPath = "output.md";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Configure Markdown save options
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Enable column alignment using space padding
            AlignColumnPadding = ' ',
            // Use the first row as the table header
            TableHeaderType = MarkdownTableHeaderType.FirstRow,
            // Set encoding and format strategy as needed
            Encoding = Encoding.UTF8,
            FormatStrategy = CellValueFormatStrategy.DisplayString
        };

        // Save the workbook as a Markdown document with aligned tables
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Workbook successfully converted to Markdown with aligned tables.");
    }
}