// Title: Export an Excel workbook with embedded hyperlinks to PDF while preserving clickable links using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert Excel to PDF with active hyperlinks | how to retain hyperlink functionality when saving workbook as PDF in .NET | preserve clickable URLs in PDF generated from Excel using Aspose.Cells | export Excel sheet containing links to PDF without losing hyperlink targets
// Tags: Aspose.Cells export workbook to PDF with hyperlinks | maintain hyperlink functionality in PDF export | clickable links in PDF generated from Excel | hyperlink handling in Aspose.Cells PDF output | save worksheet as PDF with active URLs

using System;
using System.IO;
using Aspose.Cells;

namespace HyperlinkPdfDemo
{
    // The example creates a workbook, adds display text and hyperlinks to cells, then saves the workbook as a PDF using Aspose.Cells for .NET, ensuring the hyperlinks remain clickable in the resulting PDF file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add display text and hyperlink for Google in cell A1
                sheet.Cells["A1"].PutValue("Google");
                sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.google.com");

                // Add display text and hyperlink for StackOverflow in cell B2
                sheet.Cells["B2"].PutValue("StackOverflow");
                sheet.Hyperlinks.Add(1, 1, 1, 1, "https://stackoverflow.com");

                // Define output PDF file path
                string pdfPath = "Hyperlinks.pdf";

                // Export the workbook to PDF format
                workbook.Save(pdfPath, SaveFormat.Pdf);

                // Verify that the PDF file was created
                if (File.Exists(pdfPath))
                {
                    Console.WriteLine($"PDF file '{pdfPath}' created successfully.");
                }
                else
                {
                    Console.WriteLine($"Failed to create PDF file '{pdfPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
