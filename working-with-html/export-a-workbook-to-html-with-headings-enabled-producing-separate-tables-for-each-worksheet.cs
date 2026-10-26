// Title: Export an Excel workbook to HTML with worksheet headings and separate tables for each sheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export workbook to HTML with each worksheet as its own table | How to include worksheet headings when saving Excel as HTML using Aspose.Cells | Save Excel file as HTML with separate tables per sheet in .NET | C# HtmlSaveOptions to generate HTML with sheet headers using Aspose.Cells | Export multiple worksheets to HTML with headings using Aspose.Cells library
// Tags: Aspose.Cells HtmlSaveOptions export multiple worksheets to HTML | C# export Excel workbook to HTML with sheet headings | Aspose.Cells generate separate HTML tables per worksheet | save workbook as HTML with worksheet headers using .NET | Aspose.Cells HTML export with distinct sheet tables

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, uses HtmlSaveOptions (which include worksheet headings by default) and saves the workbook as a single HTML file where each worksheet appears as a separate table, handling any runtime errors gracefully.
class ExportWorkbookToHtml
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (defaults already export headers)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to HTML: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime errors and display a friendly message
            Console.WriteLine($"An error occurred during export: {ex.Message}");
        }
    }
}
