// Title: Load an Excel workbook from a file path and generate an HTML string with tooltips enabled using Aspose.Cells for .NET
// Common Searches: C# load Excel file and export to HTML with tooltips using Aspose.Cells | Aspose.Cells HtmlSaveOptions enable tooltip text when converting workbook to HTML | Convert Excel workbook to HTML string in memory stream with Aspose.Cells C# | Get HTML preview of an Excel workbook with tooltips via Aspose.Cells library | How to return HTML from a loaded workbook without writing to disk using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions tooltip | Excel to HTML conversion C# | Load workbook from file Aspose.Cells | MemoryStream HTML export Aspose.Cells | C# utility method workbook to HTML

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// Provides a C# utility method that checks for the Excel file, loads it into an Aspose.Cells Workbook, uses HtmlSaveOptions (tooltips enabled by default) to save the workbook to a MemoryStream as HTML, and returns the UTF‑8 HTML string, with robust error handling.
public static class WorkbookHtmlUtility
{
    /// <param name="filePath">Full path to the Excel file to load.</param>
    /// <returns>HTML string of the workbook.</returns>
    public static string LoadWorkbookAndGetHtml(string filePath)
    {
        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"The file '{filePath}' was not found.");

        try
        {
            // Load the workbook from the given file path
            var workbook = new Workbook(filePath);

            // Configure HTML save options (tooltips are enabled by default in recent versions)
            var htmlOptions = new HtmlSaveOptions();

            // Save the workbook to a memory stream in HTML format
            using var memoryStream = new MemoryStream();
            workbook.Save(memoryStream, htmlOptions);

            // Convert the memory stream contents to a UTF-8 string
            return Encoding.UTF8.GetString(memoryStream.ToArray());
        }
        catch (Exception ex)
        {
            // Wrap any exception to provide clearer context
            throw new InvalidOperationException("An error occurred while converting the workbook to HTML.", ex);
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: dotnet run <ExcelFilePath>");
                return;
            }

            string excelPath = args[0];

            // Convert workbook to HTML
            string html = WorkbookHtmlUtility.LoadWorkbookAndGetHtml(excelPath);

            // Output a short preview of the HTML
            Console.WriteLine("HTML conversion successful. Preview:");
            Console.WriteLine(html.Substring(0, Math.Min(500, html.Length)));
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
