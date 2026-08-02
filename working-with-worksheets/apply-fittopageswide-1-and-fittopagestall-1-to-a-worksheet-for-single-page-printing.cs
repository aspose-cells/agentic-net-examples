// Title: Aspose.Cells C# – Fit a worksheet to a single printed page (FitToPagesWide=1, FitToPagesTall=1)
// Description: Creates an in‑memory workbook, adds sample data, configures the worksheet's PageSetup so the entire sheet scales to one page, and saves the file. Demonstrates how to use FitToPagesWide and FitToPagesTall for single‑page printing in Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# page setup | FitToPagesWide | FitToPagesTall | single page print Aspose.Cells | .NET Excel scaling | print worksheet on one page | Aspose.Cells example PDF export
// Common Searches: Aspose.Cells set worksheet to one page C# | FitToPagesWide=1 FitToPagesTall=1 example | how to print Excel sheet on a single page using Aspose.Cells | C# page setup fit to one page Aspose.Cells | scale worksheet to one page .NET
// Developer Intent: Configure a worksheet so that all rows and columns are printed on a single page by setting FitToPagesWide and FitToPagesTall to 1.
// Use Cases: Generate a compact PDF report where the whole sheet must appear on one page. | Create printable invoices that automatically fit on a single sheet of paper. | Produce a summary dashboard that prints without pagination for easy distribution.
// AI Prompts: Show C# code that sets FitToPagesWide and FitToPagesTall to 1 and exports the workbook to PDF with Aspose.Cells. | Give an example that combines page orientation, margins, and FitToPages settings to achieve one‑page output. | Explain how to programmatically confirm that the configured page setup will result in a single printed page.

using System;
using Aspose.Cells;

namespace AsposeCellsFitToSinglePage
{
    // Creates an in‑memory workbook, adds sample data, configures the worksheet's PageSetup so the entire sheet scales to one page, and saves the file. Demonstrates how to use FitToPagesWide and FitToPagesTall for single‑page printing in Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some data so the page has content
            worksheet.Cells["A1"].PutValue("Sample Data 1");
            worksheet.Cells["B2"].PutValue("Sample Data 2");

            // Configure page setup to fit the entire sheet on one page
            worksheet.PageSetup.FitToPagesWide = 1; // one page wide
            worksheet.PageSetup.FitToPagesTall = 1; // one page tall

            // Save the workbook to a file
            workbook.Save("FitToSinglePage.xlsx");

            Console.WriteLine("Workbook saved with FitToPagesWide=1 and FitToPagesTall=1.");
        }
    }
}
