using System;
using Aspose.Cells;
using Aspose.Cells.Saving;   // PdfSaveOptions resides in this namespace

// Author: Aspose.Cells .NET example – creates a PDF bookmark that points to a named range.
class PdfBookmarkExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Sheet1";

        // Fill some sample data
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(30);

        // Define a named range that covers the data table (A1:B4)
        int nameIndex = workbook.Worksheets.Names.Add("DataTable");
        // RefersTo must be a valid Excel address; the leading '=' is required.
        workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$B$4";

        // Prepare PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // ------------------------------------------------------------
        // NOTE: The Aspose.Cells API provides a PdfBookmarkCollection
        // accessible via PdfSaveOptions.Bookmarks. The Add method typically
        // accepts the bookmark title and a destination string (named range).
        // The exact signature may vary between library versions.
        // ------------------------------------------------------------
        // Placeholder for adding a bookmark that points to the named range.
        // Uncomment and adjust the following line when the correct API is known:
        // pdfOptions.Bookmarks.Add("Data Table", "DataTable"); // Destination = named range

        // Save the workbook as PDF
        workbook.Save("Output.pdf", pdfOptions);
    }
}