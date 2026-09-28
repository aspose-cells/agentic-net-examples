// Title: Convert HTML containing CSS gradient backgrounds to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an HTML file with CSS gradient background definitions into an Aspose.Cells Workbook using LoadOptions(LoadFormat.Html) and saves it as a PDF while retaining the gradient styling. | Show how to configure Aspose.Cells to preserve CSS gradient fills when converting an HTML document to PDF in a .NET application.
// Common Searches: how to keep CSS gradient backgrounds when converting HTML to PDF using Aspose.Cells C# | Aspose.Cells HTML to PDF conversion preserving gradient fills example | C# load HTML with gradient backgrounds into workbook and export to PDF Aspose.Cells
// Tags: Aspose.Cells HTML to PDF conversion with CSS gradients | LoadOptions Html format workbook import | SaveFormat.Pdf preserving stylesheet rendering | C# gradient background export using Aspose.Cells | PDF generation from HTML preserving visual styles

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example verifies the input HTML file, loads it into an Aspose.Cells Workbook with LoadOptions set to Html format, and then saves the workbook as a PDF using SaveFormat.Pdf, ensuring that CSS gradient backgrounds are retained in the resulting document.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.html";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input HTML file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the HTML content into a workbook using Html load options
                var loadOptions = new LoadOptions(LoadFormat.Html);
                var workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
