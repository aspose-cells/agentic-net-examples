// Title: Aspose.Cells .NET: Place a Custom Footer at (0,0) on Every PDF Page
// Description: Shows how to set the footer margin to zero inches, add custom footer text, and export a workbook to PDF so the footer is anchored at the page origin (0,0) on each page.
// Keywords: Aspose.Cells PDF footer | footer margin zero inches | custom footer position | C# Aspose.Cells example | set footer coordinates | page origin footer | .NET workbook to PDF
// Common Searches: Aspose.Cells set footer margin to zero | place footer at (0,0) in PDF using C# | add custom footer to every PDF page Aspose.Cells | how to position footer at bottom left in Aspose.Cells PDF | Aspose.Cells PDF footer placement tutorial
// Developer Intent: Add a custom footer that starts at coordinate (0,0) on each page of a PDF generated from an Aspose.Cells workbook.
// Use Cases: Insert a legal disclaimer that must begin at the very edge of every PDF page. | Generate invoices where terms and conditions appear flush with the page origin. | Create multi‑page reports that require a precisely positioned footer for branding consistency.
// AI Prompts: Provide C# code to set the footer margin to zero inches and add custom text using Aspose.Cells for .NET. | Explain how to anchor a footer at (0,0) on each PDF page and save the workbook as PDF. | Show how to switch the footer section (left, center, right) while keeping it positioned at the page origin.

using System;
using Aspose.Cells;

// Shows how to set the footer margin to zero inches, add custom footer text, and export a workbook to PDF so the footer is anchored at the page origin (0,0) on each page.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample content (optional)
        worksheet.Cells["A1"].PutValue("Sample content for PDF");

        // Position the footer at the very bottom-left corner of the page
        // Set the footer margin to zero inches so the footer starts at coordinate (0,0)
        worksheet.PageSetup.FooterMarginInch = 0;

        // Set the custom footer text in the center section (section index 1)
        // You can change the section index to 0 (left) or 2 (right) as needed
        worksheet.PageSetup.SetFooter(1, "Custom Footer Text");

        // Save the workbook as PDF; the footer will appear on every page at (0,0)
        workbook.Save("CustomFooter.pdf", SaveFormat.Pdf);
    }
}
