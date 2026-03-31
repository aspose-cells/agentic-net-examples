using System;
using Aspose.Cells;

class HeaderFooterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        PageSetup pageSetup = sheet.PageSetup;

        // Enable different headers/footers for the first page and for odd/even pages
        pageSetup.IsHFDiffFirst = true;
        pageSetup.IsHFDiffOddEven = true;

        // ---------- Odd pages (standard) ----------
        // Header: Left = file name, Center = custom title with font, Right = current date
        pageSetup.SetHeader(0, "&F"); // Left section
        pageSetup.SetHeader(1, "&\"Arial,Bold\"&12&CHello Report"); // Center section with font and size
        pageSetup.SetHeader(2, "&D"); // Right section

        // Footer: Left = static text, Center = page number of total pages, Right = current time
        pageSetup.SetFooter(0, "Confidential");
        pageSetup.SetFooter(1, "Page &P of &N");
        pageSetup.SetFooter(2, "&T");

        // ---------- Even pages ----------
        // Header: custom text with different fonts and colors
        pageSetup.SetEvenHeader(0, "&\"Calibri,Italic\"&10Even Left");
        pageSetup.SetEvenHeader(1, "&C&\"Calibri,Bold\"&12Even Center");
        pageSetup.SetEvenHeader(2, "&R&KFF0000Even Right"); // Red color

        // Footer: custom text and page numbers
        pageSetup.SetEvenFooter(0, "Even Footer Left");
        pageSetup.SetEvenFooter(1, "Page &P of &N");
        pageSetup.SetEvenFooter(2, "&\"Times New Roman\"&8Even Right");

        // ---------- First page ----------
        // Header: distinct content for each section
        pageSetup.SetFirstPageHeader(0, "&\"Verdana\"&14First Page Left");
        pageSetup.SetFirstPageHeader(1, "&CFirst Page Center");
        pageSetup.SetFirstPageHeader(2, "&RFirst Page Right");

        // Footer: distinct content for each section
        pageSetup.SetFirstPageFooter(0, "First Page Footer Left");
        pageSetup.SetFirstPageFooter(1, "Page &P");
        pageSetup.SetFirstPageFooter(2, "Generated on &D");

        // Save the workbook with the configured headers and footers
        workbook.Save("HeaderFooterConfigured.xlsx");
    }
}