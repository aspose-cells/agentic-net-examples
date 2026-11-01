// Title: Convert HTML containing embedded SVG to Excel with scalable vector shapes using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an HTML file with inline SVG using Aspose.Cells HtmlLoadOptions and saves it as an .xlsx workbook while keeping the SVG as vector shapes. | Show how to check for the source HTML file, create the destination folder, and perform an HTML‑to‑Excel conversion that retains SVG graphics in a .NET console application. | Explain which Aspose.Cells settings are required to import SVG elements as editable shapes during HTML import.
// Common Searches: Aspose.Cells HTML to Excel conversion preserve SVG vector shapes .NET | C# load HTML page with embedded SVG and export to XLSX keeping graphics | How to keep inline SVG images when converting HTML to Excel with Aspose.Cells | HtmlLoadOptions SVG support example Aspose.Cells C#
// Tags: HTML to Excel conversion with SVG preservation Aspose.Cells | Aspose.Cells HtmlLoadOptions SVG graphics handling | C# load HTML containing SVG and save as XLSX | retain scalable vector shapes in Excel output | error handling for missing HTML file Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// // Loads an HTML file that may include embedded SVG graphics using Aspose.Cells HtmlLoadOptions, converts the content to an Excel workbook, ensures the output directory exists, and saves the result as an .xlsx file while preserving the SVG images as scalable vector shapes.
class Program
{
    static void Main()
    {
        // Path to the source HTML file that contains embedded SVG graphics
        string htmlPath = "input.html";

        // Desired output Excel file path
        string excelPath = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The HTML file \"{htmlPath}\" was not found.");
                return;
            }

            // Configure load options for HTML (default settings)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML content (including SVG) into a new workbook
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Successfully converted \"{htmlPath}\" to \"{excelPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
