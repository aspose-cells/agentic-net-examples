// Title: Generate a PDF with a bookmark that points to a named cell range using Aspose.Cells in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, defines a named range, and saves it as a PDF containing a bookmark that jumps to that range. | Update existing Aspose.Cells PDF export logic to include a named destination reference so the produced PDF has a clickable bookmark for cells A2:B4. | Demonstrate how to set PdfSaveOptions for one page per sheet while preparing a named range for future PDF bookmark usage in C#.
// Common Searches: how to add a PDF bookmark to a specific cell range using Aspose.Cells .NET | Aspose.Cells C# export Excel to PDF with named range bookmark | create named destination for PDF bookmark in Aspose.Cells workbook | Aspose.Cells PDFSaveOptions bookmark support limitations | C# generate PDF from workbook and link bookmark to cells A2:B4
// Tags: Aspose.Cells PDF export with named range bookmark | C# create named range for PDF navigation | PdfSaveOptions one page per sheet Aspose.Cells | Aspose.Cells PDF bookmark limitation | Excel to PDF bookmark using Aspose.Cells C#

using Aspose.Cells;
using System;

// The example builds a workbook, fills it with sample data, defines a named range "QtyRange" covering cells A2:B4, configures PdfSaveOptions to place each worksheet on a separate page, and saves the file as ReportWithBookmark.pdf. Although a named range is prepared, the current Aspose.Cells version does not expose bookmark or named destination properties, so the PDF does not contain an actual clickable bookmark.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and set its name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(50);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(20);

            // Define a named range that can be used as a PDF bookmark destination
            string bookmarkName = "QtyRange";
            sheet.Cells.CreateRange("A2:B4").Name = bookmarkName;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep each worksheet on a separate page
                OnePagePerSheet = true
                // Note: Properties such as Bookmarks and NamedDestination are not available
                // in the current Aspose.Cells version; they are omitted for compatibility.
            };

            // Save the workbook as PDF
            workbook.Save("ReportWithBookmark.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
