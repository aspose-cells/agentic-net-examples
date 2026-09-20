// Title: Generate a PDF report listing each worksheet’s original and updated paper size using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook, records the PaperSize of every worksheet, changes each sheet to A4, and creates a PDF containing a table of worksheet name, original size, and new size. | Extend the program to also capture each worksheet’s orientation (portrait or landscape) and add that column to the PDF report. | Add comprehensive error handling so that if a worksheet’s page setup cannot be modified, the issue is logged and the PDF generation continues.
// Common Searches: Aspose.Cells .NET export worksheet paper size changes to a PDF report | C# list Excel sheet page setup properties before and after modification | Create PDF summary of Excel worksheets paper size using Aspose.Cells | Record original and new paper size of each sheet when converting to PDF with Aspose.Cells | Generate a table of worksheet names and paper sizes in a PDF using Aspose.Cells for .NET
// Tags: Aspose.Cells export worksheet paper size to PDF | C# modify worksheet PageSetup PaperSize | generate PDF report from Excel with Aspose.Cells | record before and after page setup properties | auto-fit columns before saving PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads source.xlsx, iterates through each worksheet to capture its original PaperSize, sets the size to A4, writes the worksheet name with before/after sizes into a new workbook, auto‑fits columns, and saves the result as PaperSizeReport.pdf.
class PdfReportGenerator
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string reportPdfPath = "PaperSizeReport.pdf";

            // Verify source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook;
            try
            {
                sourceWorkbook = new Workbook(sourcePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load source workbook: {ex.Message}");
                return;
            }

            // Create a new workbook for the PDF report
            Workbook reportWorkbook = new Workbook();
            Worksheet reportSheet = reportWorkbook.Worksheets[0];
            reportSheet.Name = "PaperSizeReport";

            // Write table headers
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Paper Size Before");
            reportSheet.Cells["C1"].PutValue("Paper Size After");

            int reportRow = 1; // zero‑based index; row 1 is the second row

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet ws in sourceWorkbook.Worksheets)
            {
                // Capture paper size before modification
                string beforeSize = ws.PageSetup.PaperSize.ToString();

                // Modify the paper size (example: set to A4)
                ws.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Capture paper size after modification
                string afterSize = ws.PageSetup.PaperSize.ToString();

                // Write data into the report sheet
                reportSheet.Cells[reportRow, 0].PutValue(ws.Name);      // Column A
                reportSheet.Cells[reportRow, 1].PutValue(beforeSize); // Column B
                reportSheet.Cells[reportRow, 2].PutValue(afterSize);  // Column C

                reportRow++;
            }

            // Auto‑fit columns for better appearance
            reportSheet.AutoFitColumns();

            // Save the report workbook as a PDF file
            try
            {
                reportWorkbook.Save(reportPdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Report saved to {reportPdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save PDF report: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
