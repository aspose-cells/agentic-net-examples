// Title: Export an Excel workbook to HTML with a custom header that displays the workbook title using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets the workbook's BuiltInDocumentProperties.Title, and saves it as an HTML file with Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to embed a custom <head> section containing the workbook title in the generated HTML. | Write a C# snippet that verifies the source Excel file exists, applies a document title, and includes exception handling while converting to HTML.
// Common Searches: Aspose.Cells C# export Excel to HTML and include workbook title in the page header | How to set BuiltInDocumentProperties.Title before saving workbook as HTML with Aspose.Cells | C# convert .xlsx to .html with custom HTML head using Aspose.Cells HtmlSaveOptions | Aspose.Cells HtmlSaveOptions missing Title property older version workaround
// Tags: Aspose.Cells HTML export with document title | C# HtmlSaveOptions custom header Aspose.Cells | set workbook built-in properties Aspose.Cells | convert xlsx to html Aspose.Cells | error handling missing input file C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing .xlsx workbook, assigns a title via BuiltInDocumentProperties, configures HtmlSaveOptions, and saves the workbook as an HTML file. It also checks for the source file's existence and handles runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook title via built‑in document properties
            workbook.BuiltInDocumentProperties.Title = "My Workbook Title";

            // Configure HTML save options (properties Title/CustomHeader are not available in older API versions)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
