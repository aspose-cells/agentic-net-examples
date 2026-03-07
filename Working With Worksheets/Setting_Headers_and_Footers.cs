using System;
using Aspose.Cells;

namespace AsposeCellsHeaderFooterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for header/footer configuration
            PageSetup pageSetup = worksheet.PageSetup;

            // -------------------------------------------------
            // Set standard headers (left, center, right sections)
            // -------------------------------------------------
            // Left section: file name
            pageSetup.SetHeader(0, "&F");
            // Center section: page number of total pages
            pageSetup.SetHeader(1, "Page &P of &N");
            // Right section: current date
            pageSetup.SetHeader(2, "&D");

            // -------------------------------------------------
            // Set standard footers (left, center, right sections)
            // -------------------------------------------------
            // Left section: custom text
            pageSetup.SetFooter(0, "Confidential");
            // Center section: sheet name
            pageSetup.SetFooter(1, "&A");
            // Right section: current time
            pageSetup.SetFooter(2, "&T");

            // -------------------------------------------------
            // Enable different headers/footers for odd and even pages
            // -------------------------------------------------
            pageSetup.IsHFDiffOddEven = true;
            // Even page header: different left text
            pageSetup.SetEvenHeader(0, "Even Page Header");
            // Even page footer: different left text
            pageSetup.SetEvenFooter(0, "Even Page Footer");

            // -------------------------------------------------
            // Enable different header/footer for the first page
            // -------------------------------------------------
            pageSetup.IsHFDiffFirst = true;
            // First page header: custom left text
            pageSetup.SetFirstPageHeader(0, "First Page Header");
            // First page footer: custom left text
            pageSetup.SetFirstPageFooter(0, "First Page Footer");

            // -------------------------------------------------
            // Save the workbook to an XLSX file
            // -------------------------------------------------
            workbook.Save("HeadersFootersDemo.xlsx");

            // -------------------------------------------------
            // Optional: demonstrate clearing all headers and footers
            // -------------------------------------------------
            pageSetup.ClearHeaderFooter();
            workbook.Save("HeadersFootersCleared.xlsx");
        }
    }
}