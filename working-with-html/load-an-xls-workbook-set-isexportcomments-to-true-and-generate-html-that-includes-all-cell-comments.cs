// Title: Export all Excel cell comments to HTML with Aspose.Cells for .NET
// Common Searches: c# asp.net export excel cell comments to html using aspose.cells | how to include cell comments when saving an xls workbook as html with aspose.cells | set IsExportComments true in HtmlSaveOptions example c# | convert legacy xls file to html preserving comments using aspose.cells | asp.net generate html from workbook with comments using aspose.cells library
// Tags: cell comment export via HtmlSaveOptions | Aspose.Cells HTML export with comments | IsExportComments option usage | XLS workbook to HTML preserving annotations | C# Aspose.Cells generate HTML with comments

using System;
using Aspose.Cells;

// Loads an XLS workbook, enables the IsExportComments flag in HtmlSaveOptions to include every cell comment, and saves the workbook as an HTML file.
class Program
{
    static void Main()
    {
        // Path to the source XLS workbook
        string inputPath = "input.xls";

        // Path where the generated HTML will be saved
        string outputPath = "output.html";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options to include all cell comments
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export comments for all cells
            IsExportComments = true
        };

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("HTML file generated with comments exported.");
    }
}
