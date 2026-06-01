using System;
using Aspose.Cells;

class PreserveHyperlinksPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put display text into a cell
        worksheet.Cells["A1"].PutValue("Visit Aspose");

        // Add a hyperlink to the cell (A1)
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.aspose.com");

        // Save the workbook as PDF; hyperlinks remain clickable in the PDF output
        workbook.Save("HyperlinksPreserved.pdf", SaveFormat.Pdf);
    }
}