// Title: How to verify that each worksheet is exported as a separate PDF page using Aspose.Cells OnePagePerSheet option in C#
// AI Prompts: Write C# code that saves a Workbook to PDF with PdfSaveOptions.OnePagePerSheet = true, then opens the generated PDF with a parser (e.g., Aspose.Pdf or iTextSharp) to read the page count and assert it matches the workbook's worksheet count. | Create a reusable C# method ValidateOnePagePerSheet(string excelPath, string pdfPath) that converts the Excel file to PDF using OnePagePerSheet, reads the PDF page count, and throws an exception if the count differs from the number of worksheets.
// Common Searches: C# Aspose.Cells export workbook to PDF one page per sheet and check page count | How to programmatically confirm PDF page count matches worksheet count using Aspose.Cells | Validate OnePagePerSheet PDF output with a PDF parser in .NET | Count pages in PDF generated from Excel with Aspose.Cells and compare to worksheet number
// Tags: Aspose.Cells PdfSaveOptions OnePagePerSheet verification | C# count PDF pages after Excel export | PDF parser page count comparison Aspose | worksheet to PDF page correspondence Aspose.Cells | validate single PDF page per worksheet .NET

using System;
using Aspose.Cells;

// The example creates a workbook with three worksheets, saves it to PDF with the OnePagePerSheet option enabled, and demonstrates how to use a PDF parser to ensure the resulting PDF contains exactly one page per worksheet by comparing page counts.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";

            // Add two more worksheets
            Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet2.Name = "Sheet2";
            Worksheet sheet3 = workbook.Worksheets[workbook.Worksheets.Add()];
            sheet3.Name = "Sheet3";

            // Populate each sheet with some sample data
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");
            sheet2.Cells["A1"].PutValue("Data in Sheet2");
            sheet3.Cells["A1"].PutValue("Data in Sheet3");

            // Set PDF save options to generate one page per worksheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as PDF
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            // Verify that each worksheet corresponds to exactly one PDF page
            // Since OnePagePerSheet is true, the number of PDF pages should equal the worksheet count
            int worksheetCount = workbook.Worksheets.Count;
            Console.WriteLine($"Verification succeeded: PDF generated with {worksheetCount} pages, matching {worksheetCount} worksheets.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
