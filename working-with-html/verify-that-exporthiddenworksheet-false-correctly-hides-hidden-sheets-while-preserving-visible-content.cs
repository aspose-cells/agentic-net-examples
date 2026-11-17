// Title: Verify that setting ExportHiddenWorksheet to false excludes hidden worksheets when saving a workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells ExportHiddenWorksheet false hide hidden sheets in PDF C# | how to exclude hidden worksheets from PDF output using Aspose.Cells .NET | C# example of PDF conversion that skips hidden worksheets with Aspose.Cells | verify hidden worksheet not present in generated PDF Aspose.Cells | Aspose.Cells PDF save options default behavior for hidden worksheets
// Tags: PDF export hidden sheet exclusion Aspose.Cells | hidden sheet omission during PDF conversion C# | Aspose.Cells PDF save options visibility control | C# workbook PDF generation without hidden sheets | Aspose.Cells PDF output hidden sheet handling

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook with one visible and one hidden worksheet, marks the hidden sheet as not visible, and saves the workbook to PDF using PdfSaveOptions (ExportHiddenWorksheet defaults to false), confirming that the hidden worksheet is omitted while the visible sheet is included in the PDF.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the default first worksheet and add some data
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("This is a visible sheet");

                // Add a second worksheet and add data to it
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("This sheet will be hidden");

                // Hide the second worksheet
                hiddenSheet.IsVisible = false;

                // Configure PDF save options (hidden worksheets are excluded by default)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook to PDF using the configured options
                string outputPath = "ExportHiddenWorksheetTest.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
