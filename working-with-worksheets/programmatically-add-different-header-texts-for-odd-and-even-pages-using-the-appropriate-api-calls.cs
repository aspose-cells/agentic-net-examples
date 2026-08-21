// Title: Set Different Odd and Even Page Headers in Excel with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, enables distinct odd/even page headers via PageSetup.IsHFDiffOddEven, assigns a centered header for odd pages with SetHeader and a centered header for even pages with SetEvenHeader, and saves the workbook as DifferentOddEvenHeaders.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel header | odd page header | even page header | PageSetup | IsHFDiffOddEven | SetHeader | SetEvenHeader | sample code | GitHub example | programmatic header
// Common Searches: Aspose.Cells set odd and even page headers C# | PageSetup.IsHFDiffOddEven example .NET | How to add different headers for odd/even pages in Excel using Aspose.Cells | SetEvenHeader Aspose.Cells tutorial | C# code to configure odd/even headers in Excel workbook
// Developer Intent: Generate an Excel workbook where odd and even printed pages display separate header texts using Aspose.Cells APIs.
// Use Cases: Print‑ready reports that require a company logo on odd pages and contact details on even pages. | Multi‑page invoices where the title appears on odd pages and the address on even pages. | Booklet creation with left‑hand (even) and right‑hand (odd) page headers for professional publishing.
// AI Prompts: Write C# code with Aspose.Cells to set a left‑aligned odd page header and a right‑aligned even page header, then export the workbook to PDF. | Explain the purpose of PageSetup.IsHFDiffOddEven and show how to switch back to a single header for all pages. | Provide an Aspose.Cells example that adds page numbers to odd page headers and the current date to even page headers.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderExample
{
    // This C# example creates a new Workbook, enables distinct odd/even page headers via PageSetup.IsHFDiffOddEven, assigns a centered header for odd pages with SetHeader and a centered header for even pages with SetEvenHeader, and saves the workbook as DifferentOddEvenHeaders.xlsx.
    public class DifferentOddEvenHeaders
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the PageSetup object of the worksheet
                PageSetup pageSetup = worksheet.PageSetup;

                // Enable different headers for odd and even pages
                pageSetup.IsHFDiffOddEven = true;

                // Set header for odd pages (default header)
                // Section 0 = Left, 1 = Center, 2 = Right
                pageSetup.SetHeader(1, "Odd Page Header - Center Section");

                // Set header for even pages
                pageSetup.SetEvenHeader(1, "Even Page Header - Center Section");

                // Define output file path
                string outputPath = "DifferentOddEvenHeaders.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DifferentOddEvenHeaders.Run();
        }
    }
}
