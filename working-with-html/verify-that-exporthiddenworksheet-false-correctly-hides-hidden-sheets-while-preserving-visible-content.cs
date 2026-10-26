// Title: Export a workbook to PDF in C# while automatically omitting hidden worksheets with Aspose.Cells
// Common Searches: C# Aspose.Cells export PDF without hidden worksheets | how to hide a worksheet from PDF output using Aspose.Cells .NET | Aspose.Cells PdfSaveOptions exclude hidden sheets example | verify that hidden worksheets are not included in generated PDF Aspose.Cells | Aspose.Cells export visible sheets only PDF C# code sample
// Tags: Aspose.Cells PDF export hide worksheets | PdfSaveOptions exclude hidden sheets | C# workbook to PDF without hidden worksheets | Aspose.Cells hidden worksheet handling PDF | OnePagePerSheet PDFSaveOptions Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample creates a workbook with one visible and one hidden worksheet, marks the hidden sheet as not visible, configures PdfSaveOptions (including OnePagePerSheet), and saves the workbook to PDF, demonstrating that hidden worksheets are automatically excluded from the PDF output.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();

            // Access the default first worksheet (visible)
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");
            visibleSheet.Cells["A2"].PutValue(123);

            // Add a second worksheet (will be hidden)
            int hiddenSheetIndex = workbook.Worksheets.Add();
            Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
            hiddenSheet.Name = "HiddenSheet";
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.Cells["A2"].PutValue(456);
            hiddenSheet.IsVisible = false; // Hide the worksheet

            // Set PDF save options (hidden worksheets will be excluded by default in recent versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // each sheet gets its own page
            };

            // Save the workbook to PDF
            string pdfPath = "ExportHiddenWorksheetDemo.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("Workbook saved to PDF with hidden worksheets excluded.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
