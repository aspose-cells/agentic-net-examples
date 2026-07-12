using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class PdfBookmarkDemo
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "TargetSheet";

        // Define the cell that will serve as the bookmark destination
        sheet.Cells["B2"].PutValue("Bookmark Destination");

        // Create a PDF bookmark entry and assign a stable destination name
        PdfBookmarkEntry bookmark = new PdfBookmarkEntry
        {
            Text = "Target Sheet Bookmark",
            Destination = sheet.Cells["B2"],
            DestinationName = "StableDestName", // stable named destination
            IsOpen = true
        };

        // Set up PDF save options with the bookmark structure
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Bookmark = bookmark
        };

        // Save the workbook as a PDF file with the custom bookmark
        workbook.Save("TargetSheetBookmark.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells example – creates a PDF bookmark with a stable destination name.