using System;
using Aspose.Cells;

namespace AsposeCellsHeaderFooterDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the PageSetup object
            PageSetup pageSetup = worksheet.PageSetup;

            // Enable different headers/footers for the first page and for odd/even pages
            pageSetup.IsHFDiffFirst = true;
            pageSetup.IsHFDiffOddEven = true;

            // ----- Headers -----
            // Standard (odd) page header: left, center, right sections
            pageSetup.SetHeader(0, "&LStandard Left Header");
            pageSetup.SetHeader(1, "&CStandard Center Header");
            pageSetup.SetHeader(2, "&RStandard Right Header");

            // Even page header
            pageSetup.SetEvenHeader(0, "&LEven Left Header");
            pageSetup.SetEvenHeader(1, "&CEven Center Header");
            pageSetup.SetEvenHeader(2, "&REven Right Header");

            // First page header
            pageSetup.SetFirstPageHeader(0, "&LFirst Page Left Header");
            pageSetup.SetFirstPageHeader(1, "&CFirst Page Center Header");
            pageSetup.SetFirstPageHeader(2, "&RFirst Page Right Header");

            // ----- Footers -----
            // Standard (odd) page footer
            pageSetup.SetFooter(0, "&LOdd Left Footer");
            pageSetup.SetFooter(1, "&CPage &P of &N"); // center shows page number
            pageSetup.SetFooter(2, "&ROdd Right Footer");

            // Even page footer
            pageSetup.SetEvenFooter(0, "&LEven Left Footer");
            pageSetup.SetEvenFooter(1, "&CEven Center Footer");
            pageSetup.SetEvenFooter(2, "&REven Right Footer");

            // First page footer
            pageSetup.SetFirstPageFooter(0, "&LFirst Footer Left");
            pageSetup.SetFirstPageFooter(1, "&CFirst Footer Center");
            pageSetup.SetFirstPageFooter(2, "&RFirst Footer Right");

            // Save the workbook
            string filePath = "HeadersFooters_DifferentPages.xlsx";
            workbook.Save(filePath);

            // Load the workbook to verify the settings
            Workbook loadedWorkbook = new Workbook(filePath);
            PageSetup loadedSetup = loadedWorkbook.Worksheets[0].PageSetup;

            // Output verification
            Console.WriteLine("IsHFDiffFirst: " + loadedSetup.IsHFDiffFirst);
            Console.WriteLine("IsHFDiffOddEven: " + loadedSetup.IsHFDiffOddEven);
            Console.WriteLine("First Page Header (Center): " + loadedSetup.GetFirstPageHeader(1));
            Console.WriteLine("Even Page Header (Center): " + loadedSetup.GetEvenHeader(1));
            Console.WriteLine("Standard Header (Center): " + loadedSetup.GetHeader(1));
            Console.WriteLine("First Page Footer (Center): " + loadedSetup.GetFirstPageFooter(1));
            Console.WriteLine("Even Page Footer (Center): " + loadedSetup.GetEvenFooter(1));
            Console.WriteLine("Standard Footer (Center): " + loadedSetup.GetFooter(1));
        }
    }
}