using System;
using Aspose.Cells;

class UniqueHeaderFooterDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate the worksheet with enough rows to span multiple printed pages
        for (int row = 0; row < 200; row++)
        {
            worksheet.Cells[row, 0].PutValue($"Row {row + 1}");
            worksheet.Cells[row, 1].PutValue(row * 10);
        }

        // Access the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Enable distinct headers/footers for the first page and for odd/even pages
        pageSetup.IsHFDiffFirst = true;      // First page differs from others
        pageSetup.IsHFDiffOddEven = true;    // Odd pages differ from even pages

        // ----- First page (unique) -----
        pageSetup.SetFirstPageHeader(0, "&LFirst Page Header Left");
        pageSetup.SetFirstPageHeader(1, "&CFirst Page Header Center");
        pageSetup.SetFirstPageHeader(2, "&RFirst Page Header Right");

        pageSetup.SetFirstPageFooter(0, "&LFirst Page Footer Left");
        pageSetup.SetFirstPageFooter(1, "&CPage &P of &N - First Footer");
        pageSetup.SetFirstPageFooter(2, "&RFirst Page Footer Right");

        // ----- Odd pages (default) -----
        pageSetup.SetHeader(0, "&LOdd Page Header Left");
        pageSetup.SetHeader(1, "&CPage &P of &N - Odd Header");
        pageSetup.SetHeader(2, "&ROdd Page Header Right");

        pageSetup.SetFooter(0, "&LOdd Footer Left");
        pageSetup.SetFooter(1, "&CPage &P of &N - Odd Footer");
        pageSetup.SetFooter(2, "&ROdd Footer Right");

        // ----- Even pages -----
        pageSetup.SetEvenHeader(0, "&LEven Page Header Left");
        pageSetup.SetEvenHeader(1, "&CPage &P of &N - Even Header");
        pageSetup.SetEvenHeader(2, "&REven Page Header Right");

        pageSetup.SetEvenFooter(0, "&LEven Footer Left");
        pageSetup.SetEvenFooter(1, "&CPage &P of &N - Even Footer");
        pageSetup.SetEvenFooter(2, "&REven Footer Right");

        // Save the workbook with the configured headers and footers
        workbook.Save("UniqueHeaderFooter.xlsx");
    }
}