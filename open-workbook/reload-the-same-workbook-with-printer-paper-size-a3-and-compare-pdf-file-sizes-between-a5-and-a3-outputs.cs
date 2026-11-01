// Title: Reload an Excel workbook, set A5 and A3 page sizes, export to PDF, and compare file sizes with Aspose.Cells for .NET
// Common Searches: how to change worksheet page setup to A3 before saving as PDF using Aspose.Cells C# | compare PDF output file size when using A5 versus A3 paper size in Aspose.Cells | reload the same workbook to apply different page sizes for PDF export in C# | Aspose.Cells generate PDF with custom paper size and retrieve file size in bytes | C# code to set PaperSizeType.PaperA5 for all worksheets and save as PDF
// Tags: set worksheet page setup paper size Aspose.Cells | export workbook to PDF with custom paper size | reload workbook to reset page settings Aspose.Cells | compare generated PDF file sizes C# | A3 vs A5 PDF generation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, sets each worksheet's page setup to PaperA5, saves it as a PDF, then reloads the same workbook, changes the page setup to PaperA3, saves a second PDF, and finally compares the byte sizes of the two PDFs, indicating which one is larger.
class PdfPaperSizeComparison
{
    static void Main()
    {
        // Paths for the source workbook and the output PDFs
        string workbookPath = "input.xlsx";
        string pdfA5Path = "output_A5.pdf";
        string pdfA3Path = "output_A3.pdf";

        // Verify that the source workbook exists
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
            return;
        }

        try
        {
            // ---------- Generate PDF with A5 paper size ----------
            // Load the workbook
            Workbook wbA5 = new Workbook(workbookPath);
            // Set the default page setup to A5 for all worksheets
            foreach (Worksheet sheet in wbA5.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA5;
            }
            // Save as PDF
            wbA5.Save(pdfA5Path, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating A5 PDF: {ex.Message}");
            return;
        }

        try
        {
            // ---------- Reload workbook and generate PDF with A3 paper size ----------
            // Reload the same workbook to ensure original settings
            Workbook wbA3 = new Workbook(workbookPath);
            // Set the default page setup to A3 for all worksheets
            foreach (Worksheet sheet in wbA3.Worksheets)
            {
                sheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            }
            // Save as PDF
            wbA3.Save(pdfA3Path, SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error generating A3 PDF: {ex.Message}");
            return;
        }

        try
        {
            // ---------- Compare file sizes ----------
            long sizeA5 = new FileInfo(pdfA5Path).Length;
            long sizeA3 = new FileInfo(pdfA3Path).Length;

            Console.WriteLine($"A5 PDF size: {sizeA5} bytes");
            Console.WriteLine($"A3 PDF size: {sizeA3} bytes");

            if (sizeA3 > sizeA5)
            {
                Console.WriteLine("A3 PDF is larger than A5 PDF.");
            }
            else if (sizeA3 < sizeA5)
            {
                Console.WriteLine("A3 PDF is smaller than A5 PDF.");
            }
            else
            {
                Console.WriteLine("Both PDFs have the same file size.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error comparing PDF sizes: {ex.Message}");
        }
    }
}
