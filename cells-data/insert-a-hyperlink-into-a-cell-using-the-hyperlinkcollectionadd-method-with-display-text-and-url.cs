using System;
using Aspose.Cells;

class HyperlinkExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a hyperlink to cell A1.
        // Parameters: start cell, end cell, URL, display text, screen tip.
        worksheet.Hyperlinks.Add(
            "A1",          // startCellName (top‑left of the range)
            "A1",          // endCellName (bottom‑right of the range)
            "https://www.example.com", // address (URL)
            "Visit Example",           // textToDisplay
            "Open Example website");   // screenTip (optional)

        // Save the workbook to a file
        workbook.Save("HyperlinkWithDisplayText.xlsx");
    }
}