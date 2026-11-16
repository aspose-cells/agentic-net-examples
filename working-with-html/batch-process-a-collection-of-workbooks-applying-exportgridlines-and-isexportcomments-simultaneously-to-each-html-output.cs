// Title: Batch convert multiple Excel workbooks to HTML with grid lines and comments using Aspose.Cells for .NET (C#)
// Common Searches: c# batch export excel files to html with grid lines using aspose.cells | how to include cell comments when converting several workbooks to html with aspose.cells | asp.net process a list of xlsx files and save each as html preserving grid lines | aspose.cells htmlsaveoptions exportgridlines and isexportcomments example for multiple workbooks | automate conversion of multiple Excel workbooks to html in C# with comments and gridlines
// Tags: Aspose.Cells batch HTML conversion | HtmlSaveOptions ExportGridLines C# | HtmlSaveOptions IsExportComments C# | convert multiple workbooks to HTML Aspose.Cells | C# export Excel grid lines to HTML | C# include cell comments in HTML export

using System;
using System.IO;
using Aspose.Cells;

// The sample iterates over an array of Excel file paths, loads each workbook with Aspose.Cells, configures HtmlSaveOptions to enable ExportGridLines and IsExportComments, and saves every workbook as an HTML file in a target folder, while handling missing files and runtime exceptions.
class HtmlExportBatchProcessor
{
    // Entry point for demonstration
    static void Main()
    {
        // Example input workbook files
        string[] workbookFiles = new string[]
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            @"C:\Workbooks\Book3.xlsx"
        };

        // Destination folder for HTML outputs
        string outputFolder = @"C:\HtmlOutputs";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each workbook
        ProcessWorkbooks(workbookFiles, outputFolder);
    }

    /// <param name="inputFiles">Array of full paths to the source workbooks.</param>
    /// <param name="outputFolder">Folder where the HTML files will be written.</param>
    static void ProcessWorkbooks(string[] inputFiles, string outputFolder)
    {
        foreach (string filePath in inputFiles)
        {
            try
            {
                // Verify that the source file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Warning: File not found – skipping '{filePath}'.");
                    continue;
                }

                // Load the workbook from the file system
                Workbook workbook = new Workbook(filePath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Export grid lines so they appear in the HTML output
                    ExportGridLines = true,

                    // Export cell comments as HTML comment markers
                    IsExportComments = true
                };

                // Build the output HTML file name (same base name as the workbook)
                string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                string htmlFullPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlFullPath, htmlOptions);
                Console.WriteLine($"Successfully exported '{filePath}' to '{htmlFullPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors and continue processing remaining files
                Console.WriteLine($"Error processing '{filePath}': {ex.Message}");
            }
        }
    }
}
