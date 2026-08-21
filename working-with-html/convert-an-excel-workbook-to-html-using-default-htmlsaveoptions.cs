// Title: Convert Excel to HTML with Aspose.Cells (C#) – default HtmlSaveOptions
// Description: Loads an Excel workbook using Aspose.Cells, applies the default HtmlSaveOptions, and saves the file as HTML. Includes a complete C# example and a console confirmation message.
// Keywords: Aspose.Cells Excel to HTML C# | HtmlSaveOptions default | Workbook.Save HTML | export Excel as HTML | C# convert .xlsx to .html | Aspose.Cells sample code | Excel to web preview
// Common Searches: Aspose.Cells convert xlsx to html C# | default HtmlSaveOptions example | save Excel workbook as HTML without custom settings | C# code to export Excel to HTML using Aspose | how to generate HTML preview from Excel file
// Developer Intent: Create an HTML representation of an Excel workbook using Aspose.Cells with the out‑of‑the‑box HtmlSaveOptions.
// Use Cases: Provide a quick web‑ready preview of Excel reports. | Automate batch conversion of multiple spreadsheets for website publishing. | Generate HTML email bodies from Excel templates without extra configuration.
// AI Prompts: Write C# code that loads an Excel file and saves it as HTML using Aspose.Cells default HtmlSaveOptions, including error handling. | Explain how to modify HtmlSaveOptions to embed images as base64 while keeping other defaults unchanged. | Show a loop that processes a folder of Excel files, converting each to HTML with a single HtmlSaveOptions instance.

using System;
using Aspose.Cells;

namespace ExcelToHtmlConversion
{
    // Loads an Excel workbook using Aspose.Cells, applies the default HtmlSaveOptions, and saves the file as HTML. Includes a complete C# example and a console confirmation message.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (can be .xlsx, .xls, etc.)
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create HtmlSaveOptions with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Define the output HTML file path
            string outputPath = "output.html";

            // Save the workbook as HTML using the default options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
        }
    }
}
