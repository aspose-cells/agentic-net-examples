// Title: Add a custom PDF attachment with a summary report generated from workbook data using Aspose.Cells in C#
// Common Searches: how to save an Aspose.Cells workbook as PDF and also export a text summary report in C# | Aspose.Cells C# generate PDF with one page per sheet and create separate summary.txt file | C# calculate Excel formulas then export PDF with PdfSaveOptions and write summary report | using MemoryStream to write Aspose.Cells PDF to disk and save accompanying text file | Aspose.Cells PDFA1b compliance example with summary report generation in .NET
// Tags: Aspose.Cells PDF generation with PdfSaveOptions | export workbook to PDF using MemoryStream C# | generate text summary report from worksheet cells Aspose.Cells | apply PdfA1b compliance in Aspose.Cells PDF export | save PDF and separate summary file .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, populates it with sample data, calculates a total formula, builds a textual summary report, saves the workbook as a PDF using PdfSaveOptions (including PDF/A‑1b compliance) into a MemoryStream, writes the PDF to disk, and writes the summary report to a separate text file.
class Program
{
    static void Main()
    {
        try
        {
            // ------------------------------------------------------------
            // 1. Create a workbook and fill it with sample data
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            Cells cells = sheet.Cells;
            cells["A1"].PutValue("Item");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(20);
            cells["A4"].PutValue("Total");
            cells["B4"].Formula = "SUM(B2:B3)";

            // Calculate formulas so that B4 contains the total
            workbook.CalculateFormula();

            // ------------------------------------------------------------
            // 2. Build a simple textual summary report from the workbook
            // ------------------------------------------------------------
            string summaryReport = $"Summary Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                                   $"Worksheet: {sheet.Name}\n" +
                                   $"Rows with data: {sheet.Cells.MaxDataRow + 1}\n" +
                                   $"Total Quantity (B4): {cells["B4"].IntValue}\n";

            byte[] summaryBytes = Encoding.UTF8.GetBytes(summaryReport);

            // ------------------------------------------------------------
            // 3. Save the workbook as PDF into a memory stream
            // ------------------------------------------------------------
            using (MemoryStream pdfStream = new MemoryStream())
            {
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    Compliance = PdfCompliance.PdfA1b // optional, demonstrates custom PDF options
                };

                workbook.Save(pdfStream, pdfOptions);
                pdfStream.Position = 0; // reset for reading

                // ------------------------------------------------------------
                // 4. Save the PDF to disk
                // ------------------------------------------------------------
                string pdfPath = "WorkbookWithSummary.pdf";
                using (FileStream file = new FileStream(pdfPath, FileMode.Create, FileAccess.Write))
                {
                    pdfStream.CopyTo(file);
                }

                Console.WriteLine($"PDF saved successfully: {pdfPath}");
            }

            // ------------------------------------------------------------
            // 5. Save the summary report as a separate text file
            // ------------------------------------------------------------
            string summaryPath = "summary.txt";
            File.WriteAllBytes(summaryPath, summaryBytes);
            Console.WriteLine($"Summary report saved successfully: {summaryPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up resources if needed
        }
    }
}
