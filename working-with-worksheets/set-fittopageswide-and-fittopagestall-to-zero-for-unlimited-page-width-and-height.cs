// Title: Aspose.Cells C# – Set FitToPagesWide and FitToPagesTall to Unlimited (0)
// Description: Creates a workbook, accesses the first worksheet, uses PageSetup.SetFitToPages(0,0) to remove page‑fit limits, verifies the settings, and saves the file as an XLSX document.
// Keywords: Aspose.Cells PageSetup SetFitToPages | FitToPagesWide zero | FitToPagesTall zero | unlimited page scaling C# | remove page fit Aspose.Cells | print without scaling Aspose
// Common Searches: Aspose.Cells set FitToPagesWide to 0 | FitToPagesTall zero Aspose.Cells C# | disable page scaling in Excel with Aspose | how to print without page fit using Aspose.Cells | C# Aspose.Cells unlimited page width height
// Developer Intent: Configure a worksheet so that no automatic page‑fit scaling is applied by setting both FitToPagesWide and FitToPagesTall to zero.
// Use Cases: Generate reports that span any number of pages without forced scaling. | Create printable spreadsheets where column width is preserved. | Export data for users who need full‑size printing control.
// AI Prompts: Provide C# code that calls PageSetup.SetFitToPages(0,0) with Aspose.Cells and saves the workbook. | Explain how setting FitToPagesWide and FitToPagesTall to zero affects printing and layout. | Show how to read back the FitToPagesWide and FitToPagesTall values after calling SetFitToPages.

using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesDemo
{
    // Creates a workbook, accesses the first worksheet, uses PageSetup.SetFitToPages(0,0) to remove page‑fit limits, verifies the settings, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Set FitToPagesWide and FitToPagesTall to zero for unlimited page width and height
            // Using the SetFitToPages method as defined in the API
            pageSetup.SetFitToPages(0, 0);

            // Optionally, verify the settings (should both be zero)
            Console.WriteLine("FitToPagesWide: " + pageSetup.FitToPagesWide);
            Console.WriteLine("FitToPagesTall: " + pageSetup.FitToPagesTall);

            // Save the workbook to a file (lifecycle: save)
            workbook.Save("FitToPagesUnlimited.xlsx");

            // Indicate completion
            Console.WriteLine("Workbook saved with unlimited FitToPages settings.");
        }
    }
}
