// Title: Generate a PDF report with a summary sheet of clickable worksheet links using Aspose.Cells for .NET (C#)
// Common Searches: how to create a PDF report from an Excel workbook with clickable sheet links using Aspose.Cells in C# | C# add a summary worksheet with hyperlinks to each sheet before exporting to PDF with Aspose.Cells | Aspose.Cells PDFSaveOptions OnePagePerSheet example that includes navigation links between sheets | generate PDF where each Excel worksheet appears on its own page and a table of contents links to them | preserve Excel hyperlink functionality in PDF output using Aspose.Cells C#
// Tags: Aspose.Cells PDFSaveOptions OnePagePerSheet | Aspose.Cells add worksheet hyperlinks | Excel to PDF with worksheet navigation links Aspose.Cells | summary worksheet hyperlink generation C# | preserve hyperlinks when saving Excel as PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, creates (or reuses) a "Summary" sheet, writes each worksheet name with a hyperlink to its A1 cell, configures PdfSaveOptions to render each sheet on a separate PDF page, and saves the workbook as a PDF while preserving the clickable links.
class PdfReportGenerator
{
    static void Main()
    {
        // Paths for input Excel workbook and output PDF report
        string excelPath = "input.xlsx";
        string pdfPath = "report.pdf";

        // Verify that the input Excel file exists
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"Input file not found: {excelPath}");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(excelPath);

            // Ensure a summary worksheet exists to hold the hyperlinks
            Worksheet summarySheet = workbook.Worksheets["Summary"];
            if (summarySheet == null)
            {
                int index = workbook.Worksheets.Add();
                summarySheet = workbook.Worksheets[index];
                summarySheet.Name = "Summary";
            }

            // Populate the summary sheet with hyperlinks to each worksheet
            int row = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Skip the summary sheet itself
                if (sheet.Name == "Summary")
                    continue;

                // Write the worksheet name in the summary sheet
                summarySheet.Cells[row, 0].PutValue(sheet.Name);

                // Add a hyperlink that points to cell A1 of the target worksheet
                // Use overload that specifies a single-cell range (1 row, 1 column)
                summarySheet.Hyperlinks.Add(row, 0, 1, 1, $"{sheet.Name}!A1");

                row++;
            }

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Each worksheet will be rendered on a separate PDF page
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF file; hyperlinks will be preserved
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF report generated successfully at: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
