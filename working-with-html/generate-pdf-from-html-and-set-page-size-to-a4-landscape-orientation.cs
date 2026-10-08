// Title: Generate a PDF from an HTML file with A4 landscape orientation using Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that reads an HTML file, loads it into an Aspose.Cells Workbook, sets the first worksheet to A4 landscape, and saves the workbook as a PDF. | Create a reusable C# method that accepts an HTML string and returns a PDF byte array with A4 landscape page setup using Aspose.Cells. | Show how to extend the example to apply custom margins and let the caller specify the output directory when saving the PDF.
// Common Searches: aspocells c# convert html file to pdf with A4 landscape page size | how to set worksheet orientation to landscape before saving as PDF in Aspose.Cells | C# example for loading HTML into Aspose.Cells workbook and exporting to PDF | Aspose.Cells page setup A4 landscape for PDF output from HTML
// Tags: Aspose.Cells HTML to PDF conversion C# | A4 landscape page setup Aspose.Cells | Set worksheet orientation PDF Aspose.Cells | Load HTML into Aspose.Cells workbook | Export Aspose.Cells workbook as PDF

using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToPdfConversion
{
    // Loads an HTML file into an Aspose.Cells Workbook, configures the first worksheet to A4 landscape orientation, and saves the workbook as a PDF document.
    class HtmlToPdfConverter
    {
        static void Main()
        {
            try
            {
                // Path to the HTML input file
                string htmlPath = "input.html";

                // Verify that the HTML file exists
                if (!File.Exists(htmlPath))
                {
                    Console.WriteLine($"Input HTML file not found: {htmlPath}");
                    return;
                }

                // Load HTML content into a workbook using HtmlLoadOptions
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                Workbook workbook = new Workbook(htmlPath, loadOptions);

                // Configure page setup: A4 size, landscape orientation
                PageSetup pageSetup = workbook.Worksheets[0].PageSetup;
                pageSetup.PaperSize = PaperSizeType.PaperA4;
                pageSetup.Orientation = PageOrientationType.Landscape;

                // Save the workbook as a PDF file
                string pdfPath = "output.pdf";
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF successfully saved to: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
