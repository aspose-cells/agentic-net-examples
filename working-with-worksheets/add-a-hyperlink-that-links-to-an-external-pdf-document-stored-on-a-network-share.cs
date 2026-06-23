using System;
using Aspose.Cells;

class AddExternalPdfHyperlink
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the cell that will contain the hyperlink
        string cellName = "A1";

        // UNC path to the external PDF on a network share
        string pdfAddress = @"\\Server\Share\Document.pdf";

        // Add the hyperlink to the specified cell
        worksheet.Hyperlinks.Add(cellName, 1, 1, pdfAddress);

        // Set the display text for the hyperlink
        worksheet.Cells[cellName].PutValue("Open PDF");
        Hyperlink hyperlink = worksheet.Hyperlinks[0];
        hyperlink.TextToDisplay = "Open PDF";

        // Save the workbook
        workbook.Save("HyperlinkToPdf.xlsx");
    }
}