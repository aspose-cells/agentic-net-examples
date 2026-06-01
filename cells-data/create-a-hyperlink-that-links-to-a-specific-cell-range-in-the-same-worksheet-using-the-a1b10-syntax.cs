using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink in cell C1 that points to the range A1:B10 on the same sheet.
        // The address uses the "#" syntax to refer to an internal range.
        worksheet.Hyperlinks.Add("C1", 1, 1, "#A1:B10");

        // Optionally set the display text for the hyperlink.
        Hyperlink hyperlink = worksheet.Hyperlinks[0];
        hyperlink.TextToDisplay = "Go to A1:B10";

        // Save the workbook
        workbook.Save("HyperlinkToRange.xlsx");
    }
}