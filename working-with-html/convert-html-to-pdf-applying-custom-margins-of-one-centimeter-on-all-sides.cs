// Title: Generate a PDF from an HTML file with 1 cm page margins using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook, sets 1 cm (≈28.35 points) margins on all sides of the first worksheet, and saves the workbook as a PDF. | Show how to convert HTML to PDF with specific margins in Aspose.Cells, including the conversion factor from centimeters to points and applying the margins via Worksheet.PageSetup.
// Common Searches: Aspose.Cells C# convert HTML to PDF with specific page margins | set 1 cm margins when exporting HTML to PDF using Aspose.Cells | how to convert centimeters to points for page setup in Aspose.Cells | C# example for HTML to PDF conversion with custom margins Aspose.Cells
// Tags: Aspose.Cells HTML conversion with margin settings | C# set worksheet page margins in points | unit conversion for page margin values | Workbook.Save as PDF with margin configuration | page setup margin configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToPdfWithMargins
{
    // Loads an HTML file into a Workbook, applies 1 cm (≈28.35 points) margins on all sides of the first worksheet, and saves the result as a PDF.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input HTML and output PDF
            string htmlPath = "input.html";
            string pdfPath = "output.pdf";

            try
            {
                // Verify that the HTML source file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                    return;
                }

                // Load the HTML file into a new Workbook instance
                Workbook workbook = new Workbook(htmlPath);

                // Apply custom margins of 1 centimeter on all sides
                // 1 cm ≈ 28.3465 points (Aspose.Cells uses points for margins)
                const double cmToPoints = 28.3465;
                Worksheet sheet = workbook.Worksheets[0];
                sheet.PageSetup.TopMargin = cmToPoints;
                sheet.PageSetup.BottomMargin = cmToPoints;
                sheet.PageSetup.LeftMargin = cmToPoints;
                sheet.PageSetup.RightMargin = cmToPoints;

                // Save the workbook as a PDF document
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF successfully created at '{pdfPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
