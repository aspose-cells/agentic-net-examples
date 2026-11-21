// Title: Convert HTML with embedded SVG to Excel (XLSX) while preserving SVG as scalable vector shapes using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert HTML file containing SVG to XLSX preserving vector graphics | How to keep embedded SVG images as scalable shapes when importing HTML into a workbook with Aspose.Cells | Load HTML with inline SVG into Aspose.Cells workbook and save as Excel without rasterizing images | C# example for converting HTML to Excel while retaining SVG vector quality using Aspose.Cells | Preserve SVG graphics during HTML to XLSX conversion in .NET with Aspose.Cells
// Tags: HTML to XLSX conversion with SVG preservation | Aspose.Cells HtmlLoadOptions for SVG import | C# workbook creation from HTML containing vector graphics | saving Excel file with scalable SVG shapes | retain vector image quality in Excel output

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an HTML file that includes embedded SVG graphics using Aspose.Cells HtmlLoadOptions, creates a Workbook, and saves it as an XLSX file. The SVG images are kept as scalable vector shapes in the Excel output, ensuring no rasterization and preserving visual quality.
class Program
{
    static void Main()
    {
        // Path to the source HTML file that contains embedded SVG graphics
        string htmlFilePath = "input.html";

        // Path where the resulting Excel file will be saved
        string excelFilePath = "output.xlsx";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: The file '{htmlFilePath}' was not found.");
                return;
            }

            // Load options for HTML import (default options are sufficient for SVG conversion)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            // Load the HTML document into a new workbook using the specified options
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Save the workbook as an XLSX file, preserving the SVG graphics as vector shapes
            workbook.Save(excelFilePath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{excelFilePath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
