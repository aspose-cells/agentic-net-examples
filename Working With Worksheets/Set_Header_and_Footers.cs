using System;
using Aspose.Cells;

class SetHeaderFooterDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Clear any existing header/footer settings
        pageSetup.ClearHeaderFooter();

        // Set standard header sections
        // Left section: file name without path
        pageSetup.SetHeader(0, "&F");
        // Center section: page number of total pages
        pageSetup.SetHeader(1, "Page &P of &N");
        // Right section: current date
        pageSetup.SetHeader(2, "&D");

        // Set standard footer sections
        // Left section: custom text
        pageSetup.SetFooter(0, "Confidential");
        // Right section: sheet name
        pageSetup.SetFooter(2, "&A");

        // Enable different headers/footers for odd and even pages
        pageSetup.IsHFDiffOddEven = true;
        // Even page header (left section)
        pageSetup.SetEvenHeader(0, "Even Left Header");
        // Even page footer (left section)
        pageSetup.SetEvenFooter(0, "Even Left Footer");

        // Enable different header/footer for the first page
        pageSetup.IsHFDiffFirst = true;
        // First page header (center section)
        pageSetup.SetFirstPageHeader(1, "First Page Header Center");
        // First page footer (center section)
        pageSetup.SetFirstPageFooter(1, "First Page Footer Center");

        // Save the workbook with the configured headers and footers
        workbook.Save("HeaderFooterDemo.xlsx");
    }
}