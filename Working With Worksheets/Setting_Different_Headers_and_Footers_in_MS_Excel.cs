using System;
using Aspose.Cells;

namespace AsposeCellsHeaderFooterDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            PageSetup pageSetup = worksheet.PageSetup;

            // Enable different headers/footers for the first page
            pageSetup.IsHFDiffFirst = true;
            // Enable different headers/footers for odd and even pages
            pageSetup.IsHFDiffOddEven = true;

            // ----- Standard (odd) page header/footer -----
            // Left section: file name
            pageSetup.SetHeader(0, "&F");
            // Center section: page number of total pages
            pageSetup.SetHeader(1, "Page &P of &N");
            // Right section: current date
            pageSetup.SetHeader(2, "&D");

            // Left footer: sheet name
            pageSetup.SetFooter(0, "&A");
            // Center footer: current time
            pageSetup.SetFooter(1, "&T");
            // Right footer: custom text
            pageSetup.SetFooter(2, "Confidential");

            // ----- Even page header/footer -----
            // Left section: custom text for even pages
            pageSetup.SetEvenHeader(0, "Even Page Left Header");
            // Center section: custom text for even pages
            pageSetup.SetEvenHeader(1, "Even Page Center Header");
            // Right section: custom text for even pages
            pageSetup.SetEvenHeader(2, "Even Page Right Header");

            // Even page footer: simple text
            pageSetup.SetEvenFooter(0, "Even Page Footer");

            // ----- First page header/footer -----
            // First page left header
            pageSetup.SetFirstPageHeader(0, "First Page Header Left");
            // First page center header
            pageSetup.SetFirstPageHeader(1, "First Page Header Center");
            // First page right header
            pageSetup.SetFirstPageHeader(2, "First Page Header Right");

            // First page footer
            pageSetup.SetFirstPageFooter(0, "First Page Footer");

            // Save the workbook
            workbook.Save("HeadersFootersDemo.xlsx");
        }
    }
}