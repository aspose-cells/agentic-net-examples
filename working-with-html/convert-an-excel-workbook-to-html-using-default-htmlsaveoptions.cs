// Title: Convert Excel Workbook to HTML with Aspose.Cells Default HtmlSaveOptions (C#)
// Description: A C# console example that loads an .xlsx file into an Aspose.Cells Workbook, uses the default HtmlSaveOptions, and saves the workbook as an HTML file via Workbook.Save.
// Keywords: Aspose.Cells | Excel to HTML conversion | C# HtmlSaveOptions | .NET Excel export | Workbook.Save HTML | default HtmlSaveOptions | Aspose.Cells example
// Common Searches: Aspose.Cells export Excel to HTML C# | default HtmlSaveOptions example | convert .xlsx to .html using Aspose.Cells | C# save workbook as HTML | Aspose.Cells HTML conversion without custom settings
// Developer Intent: Generate an HTML file from an Excel workbook using Aspose.Cells with the default HtmlSaveOptions.
// Use Cases: Quickly preview an Excel report as HTML for web pages without custom styling. | Automate batch conversion of uploaded Excel files to HTML for email or documentation pipelines. | Create static HTML snapshots of workbook data in scheduled .NET jobs.
// AI Prompts: Provide C# code that loads an Excel file and saves it as HTML using Aspose.Cells with default HtmlSaveOptions, including proper error handling. | Explain how to modify the sample to embed images and CSS directly in the generated HTML with Aspose.Cells. | Show how to iterate over all .xlsx files in a folder and convert each to HTML using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlConversion
{
    // A C# console example that loads an .xlsx file into an Aspose.Cells Workbook, uses the default HtmlSaveOptions, and saves the workbook as an HTML file via Workbook.Save.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Path where the HTML output will be saved
            string outputPath = "output.html";

            // Load the Excel workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // Create default HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Save the workbook as HTML using the default options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook has been converted to HTML and saved to '{outputPath}'.");
        }
    }
}
