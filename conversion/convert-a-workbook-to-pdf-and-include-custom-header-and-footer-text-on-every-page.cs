// Title: Convert an Excel workbook to PDF with custom header and footer text using Aspose.Cells for .NET
// AI Prompts: Write C# code that defines left, center, and right sections for both header and footer on a worksheet, then saves the workbook as a PDF with Aspose.Cells. | Show how to embed dynamic placeholders such as file name, current date, page number, and time into worksheet headers/footers before PDF conversion using Aspose.Cells. | Demonstrate configuring PdfSaveOptions to preserve document structure while exporting a workbook that contains custom headers and footers to PDF.
// Common Searches: how to add custom header and footer in Aspose.Cells before saving as PDF | c# Aspose.Cells set worksheet page header left center right sections | export Excel to PDF with page numbers and date using Aspose.Cells .NET | include file name in PDF header when converting workbook with Aspose.Cells | Aspose.Cells PdfSaveOptions ExportDocumentStructure example
// Tags: Aspose.Cells set header footer | worksheet page setup header sections | export workbook to PDF with custom header | PdfSaveOptions document structure Aspose.Cells | dynamic placeholders header footer .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHeaderFooterPdfDemo
{
    // The program creates a workbook, adds sample data, configures left/center/right header and footer sections with static text and dynamic fields (file name, date, page number, time), sets PdfSaveOptions to export document structure, and saves the workbook as a PDF containing the custom header and footer on every page.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["A2"].PutValue(DateTime.Now.ToString());

            // Set custom header (left, center, right sections)
            // Example: Left - file name, Center - custom text, Right - date
            worksheet.PageSetup.SetHeader(0, "&F");                     // Left section: file name
            worksheet.PageSetup.SetHeader(1, "My Custom Header Text"); // Center section: custom text
            worksheet.PageSetup.SetHeader(2, "&D");                     // Right section: current date

            // Set custom footer (left, center, right sections)
            // Example: Left - page number, Center - custom text, Right - time
            worksheet.PageSetup.SetFooter(0, "Page &P of &N");          // Left section: page numbering
            worksheet.PageSetup.SetFooter(1, "My Custom Footer Text"); // Center section: custom text
            worksheet.PageSetup.SetFooter(2, "&T");                    // Right section: current time

            // Configure PDF save options if needed (e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF with the defined header/footer
            string outputPath = "WorkbookWithHeaderFooter.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF with custom header/footer at: {outputPath}");
        }
    }
}
