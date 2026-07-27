// Title: C# Aspose.Cells: Load Workbook, Enable Tooltip Text, Export to HTML String
// Description: A concise utility that validates an Excel file path, loads the workbook with Aspose.Cells, sets HtmlSaveOptions.AddTooltipText to true, streams the HTML output to memory, applies the correct encoding, and returns the full HTML markup as a string.
// Keywords: Aspose.Cells C# HTML export | AddTooltipText HtmlSaveOptions | Excel to HTML string memory stream | convert workbook to HTML without file | tooltip text in exported HTML | Aspose.Cells in‑memory conversion | C# Excel preview with comments
// Common Searches: Aspose.Cells export Excel to HTML with tooltips C# | HtmlSaveOptions AddTooltipText example | Convert Excel workbook to HTML string memory stream | C# get HTML markup from Aspose.Cells workbook | How to enable cell comment tooltips in Aspose.Cells HTML output
// Developer Intent: Provide a reusable method that reads an Excel file, activates tooltip text in the HTML conversion, and returns the generated HTML markup directly as a string.
// Use Cases: Render an interactive spreadsheet preview on a web page where cell comments appear as hover tooltips. | Generate an email body containing the HTML representation of an uploaded Excel file with comment tooltips for better context. | Expose a REST API that accepts an Excel file, converts it to HTML with tooltip support, and returns the markup to the caller.
// AI Prompts: Write a C# function using Aspose.Cells that loads an Excel file, sets HtmlSaveOptions.AddTooltipText = true, and returns the HTML string without writing to disk. | Explain the role of HtmlSaveOptions.AddTooltipText and demonstrate how to read the HTML from a MemoryStream with proper encoding. | Suggest enhancements to allow callers to specify a custom encoding and inject a CSS stylesheet into the exported HTML.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// A concise utility that validates an Excel file path, loads the workbook with Aspose.Cells, sets HtmlSaveOptions.AddTooltipText to true, streams the HTML output to memory, applies the correct encoding, and returns the full HTML markup as a string.
public static class WorkbookHtmlUtility
{
    /// <param name="excelFilePath">Full path to the source Excel file.</param>
    /// <returns>HTML representation of the workbook with tooltip text enabled.</returns>
    public static string LoadWorkbookAndExportHtml(string excelFilePath)
    {
        // Verify that the source file exists.
        if (!File.Exists(excelFilePath))
            throw new FileNotFoundException("Excel file not found.", excelFilePath);

        // Load the workbook from the given file path.
        Workbook workbook = new Workbook(excelFilePath);

        // Create HTML save options and enable tooltip text.
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            AddTooltipText = true
        };

        // Save the workbook to a memory stream using the HTML options.
        using (MemoryStream htmlStream = new MemoryStream())
        {
            workbook.Save(htmlStream, htmlOptions);

            // Ensure the stream position is at the beginning before reading.
            htmlStream.Position = 0;

            // Determine the encoding to use (default is UTF-8 if not set).
            Encoding encoding = htmlOptions.Encoding ?? Encoding.UTF8;

            // Convert the stream contents to a string and return.
            using (StreamReader reader = new StreamReader(htmlStream, encoding))
            {
                return reader.ReadToEnd();
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Determine the Excel file path: from arguments or a default placeholder.
        string excelFilePath = args.Length > 0 ? args[0] : "sample.xlsx";

        // Check file existence before proceeding.
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Error: File not found - {excelFilePath}");
            return;
        }

        try
        {
            // Export workbook to HTML with tooltip text enabled.
            string htmlContent = WorkbookHtmlUtility.LoadWorkbookAndExportHtml(excelFilePath);
            Console.WriteLine("HTML export successful. Output:");
            Console.WriteLine(htmlContent);
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
