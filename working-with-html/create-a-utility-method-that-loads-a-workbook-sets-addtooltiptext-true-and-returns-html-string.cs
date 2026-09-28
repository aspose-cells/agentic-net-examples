// Title: Generate an HTML string from an Excel workbook with tooltips using Aspose.Cells for .NET
// AI Prompts: Write a C# method that loads an .xlsx file with Aspose.Cells, turns on tooltip generation, and returns the resulting HTML markup as a string. | Build a utility that reads an Excel workbook, configures HTML export to embed cell comments as hover text, and outputs the HTML via a memory stream.
// Common Searches: Aspose.Cells export Excel to HTML with cell comments displayed as tooltips | C# convert .xlsx file to HTML string including hover tooltips | how to enable tooltip text when saving workbook as HTML using Aspose.Cells | generate HTML from Excel workbook in memory without creating a file Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions tooltip support | convert Excel workbook to HTML string C# | export workbook with cell comments as hover text | memory stream HTML conversion Aspose.Cells | load workbook from file path Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Provides a static C# utility that validates a file path, loads the Excel workbook with Aspose.Cells, sets HtmlSaveOptions.AddTooltipText to true, saves the workbook to a MemoryStream as HTML, reads the stream into a string, and returns the HTML markup; errors are wrapped in an InvalidOperationException for clearer diagnostics.
public static class WorkbookHtmlUtility
{
    /// <param name="filePath">Full path to the Excel file to be loaded.</param>
    /// <returns>HTML string of the workbook.</returns>
    public static string LoadWorkbookAndGetHtml(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path must be provided.", nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException("The specified Excel file was not found.", filePath);

        try
        {
            // Load the workbook from the given file path
            var workbook = new Workbook(filePath);

            // Configure HTML save options to include tooltip text
            var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                AddTooltipText = true
            };

            // Save the workbook to a memory stream in HTML format
            using (var htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0; // Reset stream position for reading

                // Read the HTML content from the memory stream and return it as a string
                using (var reader = new StreamReader(htmlStream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to provide context
            throw new InvalidOperationException($"Failed to convert workbook '{filePath}' to HTML.", ex);
        }
    }
}

public class Program
{
    // Simple entry point for demonstration purposes
    public static void Main(string[] args)
    {
        try
        {
            // Expect the first argument to be the Excel file path
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the full path to an Excel file as a command‑line argument.");
                return;
            }

            string excelPath = args[0];
            string html = WorkbookHtmlUtility.LoadWorkbookAndGetHtml(excelPath);
            Console.WriteLine("Generated HTML:");
            Console.WriteLine(html);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
