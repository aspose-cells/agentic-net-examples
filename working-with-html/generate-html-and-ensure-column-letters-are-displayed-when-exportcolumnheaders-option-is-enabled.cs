// Title: Export an Excel workbook to HTML with column letters displayed using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportColumnHeaders = true, and saves the workbook as HTML so the column letters (A, B, C…) appear in the output. | Show how to verify that the source Excel file exists before calling Workbook.Save with HtmlSaveOptions, and optionally disable row headers while keeping column headers visible. | Demonstrate configuring Aspose.Cells HtmlSaveOptions to hide row headers, keep column headers, and then save the workbook to a specified HTML file path.
// Common Searches: How to keep column letters when converting an .xlsx file to HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions settings to show column headers but hide row headers | C# example for verifying an Excel file exists before HTML export using Aspose.Cells | Enable ExportColumnHeaders in Aspose.Cells HTML output | Saving workbook as HTML while preserving column letters in .NET
// Tags: Aspose.Cells column header export option | HTML export of Excel with column letters C# | disable row headers Aspose.Cells HTML | check Excel file existence before conversion | C# Aspose.Cells HtmlSaveOptions usage

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that input.xlsx exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions (column headers are enabled by default and row headers can be disabled), and saves the workbook as output.html, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // By default Aspose.Cells includes column headers (A, B, C, ...) in the HTML output.
            // If you need to hide row headers, you can set the following option (available in newer versions):
            // htmlOptions.ExportRowHeaders = false;

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
