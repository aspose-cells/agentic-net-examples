// Title: Convert Arabic HTML to right‑to‑left aligned PDF using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells convert HTML containing Arabic RTL text to PDF | preserve right‑to‑left text direction when saving HTML as PDF with Aspose.Cells .NET | set default workbook style to right alignment for Arabic PDF generation in C# | load Arabic HTML into Aspose.Cells workbook and export to PDF with RTL layout | Aspose.Cells PDF export RTL Arabic alignment example C#
// Tags: aspocells html to pdf rtl conversion | c# set workbook default style right alignment | export arabic html to pdf using aspocells | right-to-left pdf generation with aspocells .net | load html into workbook aspocells c#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Arabic HTML file into an Aspose.Cells Workbook, changes the workbook's default style to right alignment to support RTL languages, and saves the workbook as a PDF, handling missing files and directory creation.
class HtmlToPdfRtlConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlPath = @"C:\Input\arabic.html";

        // Path for the output PDF file
        string pdfPath = @"C:\Output\arabic.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Input file not found: {htmlPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Load the HTML file into a Workbook
            Workbook workbook = new Workbook(htmlPath);

            // Set default style to right‑aligned (common for RTL languages)
            Style defaultStyle = workbook.DefaultStyle;
            defaultStyle.HorizontalAlignment = TextAlignmentType.Right;
            workbook.DefaultStyle = defaultStyle;

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
