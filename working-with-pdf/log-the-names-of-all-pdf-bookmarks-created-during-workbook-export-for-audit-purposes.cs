// Title: Log PDF bookmark names generated from worksheet export using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells retrieve PDF bookmark titles after saving a workbook | how to output worksheet names that become PDF bookmarks in Aspose.Cells | audit PDF bookmarks created during workbook to PDF conversion with Aspose.Cells | log PDF bookmark names during Aspose.Cells export in .NET | extract list of PDF bookmarks from exported workbook using Aspose.Cells
// Tags: Aspose.Cells export workbook to PDF with bookmarks | C# retrieve worksheet names for PDF bookmarks | log PDF bookmark names Aspose.Cells | audit PDF bookmarks during workbook conversion | PdfSaveOptions generate worksheet bookmarks

using System;
using Aspose.Cells;

// The example creates a Workbook, adds worksheets, and saves it to PDF using Aspose.Cells' PdfSaveOptions, which automatically creates a bookmark for each worksheet. After saving, it iterates through the workbook's Worksheets collection and writes each worksheet's Name to the console, providing an audit log of the PDF bookmark names.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add worksheets for demonstration
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets[0].Cells["A1"].PutValue("Summary Sheet");

            Worksheet dataSheet = workbook.Worksheets.Add("Data");
            dataSheet.Cells["A1"].PutValue("Data Sheet");

            // Configure PDF save options (Aspose.Cells automatically creates bookmarks for each worksheet)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook to PDF
            string pdfPath = "ExportedWorkbook.pdf";
            workbook.Save(pdfPath, pdfOptions);

            // Log the names of worksheets which correspond to PDF bookmarks
            Console.WriteLine("PDF Bookmarks (worksheet names):");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Bookmark Name: {sheet.Name}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
