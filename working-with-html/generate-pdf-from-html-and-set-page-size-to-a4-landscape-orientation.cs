// Title: Create a PDF from an HTML file with A4 landscape orientation using Aspose.Cells for .NET (C#)
// AI Prompts: Load an HTML file into an Aspose.Cells workbook, set the first worksheet to A4 landscape, and export it as a PDF in C#. | Show how to configure PageSetup.Orientation and PageSetup.PaperSize before saving a workbook to PDF with Aspose.Cells. | Write C# code that uses HtmlLoadOptions to read HTML, applies landscape orientation, and generates a PDF file.
// Common Searches: Aspose.Cells C# convert HTML to PDF with landscape orientation and A4 size | how to set worksheet page orientation to landscape when exporting to PDF using Aspose.Cells | specify A4 paper size for PDF output in Aspose.Cells C# example
// Tags: html to pdf conversion Aspose.Cells C# | set worksheet orientation landscape Aspose.Cells | configure A4 paper size Aspose.Cells PDF export | Workbook.Save PDF format Aspose.Cells | HtmlLoadOptions usage Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an HTML file into an Aspose.Cells Workbook, configures the first worksheet for A4 landscape layout, and saves the workbook as a PDF document.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = "input.html";

        // Path for the generated PDF file
        string pdfPath = "output.pdf";

        // Load the HTML content into a workbook
        Workbook workbook = new Workbook(htmlPath, new HtmlLoadOptions());

        // Configure page setup for the first worksheet (you can repeat for other sheets if needed)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.PageSetup.Orientation = PageOrientationType.Landscape; // Landscape orientation
        sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;           // A4 paper size

        // Save the workbook as a PDF document
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}
