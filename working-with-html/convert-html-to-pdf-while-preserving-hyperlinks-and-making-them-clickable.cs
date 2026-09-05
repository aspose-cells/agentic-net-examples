// Title: Convert HTML to PDF with clickable hyperlinks using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an HTML file into an Aspose.Cells Workbook via HtmlLoadOptions and saves it as a PDF, preserving all hyperlinks as active links. | Demonstrate how to configure the worksheet PageSetup in Aspose.Cells so the HTML content fits the width of a single PDF page before exporting. | Create a console‑application template that validates two command‑line arguments (input HTML path and output PDF path) and performs the conversion with Aspose.Cells.
// Common Searches: asp.net convert html to pdf preserving hyperlinks using aspose.cells c# | c# aspose.cells html to pdf conversion with active links | fit html content to one page when exporting to pdf with aspose.cells | command line html to pdf tool with clickable links in .net
// Tags: Aspose.Cells HTML to PDF conversion with hyperlinks | preserve clickable links in PDF using Aspose.Cells | worksheet PageSetup fit to page Aspose.Cells | HtmlLoadOptions workbook loading Aspose.Cells | C# console HTML to PDF utility Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# console application that accepts an input HTML file and an output PDF file, loads the HTML into an Aspose.Cells Workbook using HtmlLoadOptions, optionally sets each worksheet to fit the width of one page, and saves the workbook as a PDF while keeping all original hyperlinks clickable.
class HtmlToPdfConverter
{
    static void Main(string[] args)
    {
        // Validate arguments: first is input HTML file path, second is output PDF file path
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: HtmlToPdfConverter <input.html> <output.pdf>");
            return;
        }

        string htmlPath = args[0];
        string pdfPath = args[1];

        // Load the HTML file into a Workbook.
        // Aspose.Cells can parse HTML and create worksheets with the same layout.
        Workbook workbook = new Workbook(htmlPath, new HtmlLoadOptions());

        // Optional: adjust page setup if needed (e.g., fit to one page)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0; // unlimited height
        }

        // Save the workbook as PDF.
        // Hyperlinks present in the HTML are converted to cell hyperlinks and are preserved in the PDF.
        workbook.Save(pdfPath, SaveFormat.Pdf);

        Console.WriteLine($"HTML file '{htmlPath}' has been successfully converted to PDF '{pdfPath}'.");
    }
}
