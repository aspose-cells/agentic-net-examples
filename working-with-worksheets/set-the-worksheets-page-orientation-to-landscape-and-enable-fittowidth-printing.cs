// Title: Aspose.Cells for .NET: Set Worksheet Landscape Orientation & Fit‑to‑Width Printing (C#)
// Description: Shows how to use Aspose.Cells in C# to change a worksheet’s page orientation to landscape and configure PageSetup so all columns fit on one printed page width (FitToPagesWide = 1, FitToPagesTall = 0). The workbook is saved as LandscapeFitToWidth.xlsx.
// Keywords: Aspose.Cells | C# page orientation | landscape printing | fit to width | FitToPagesWide | FitToPagesTall | PageSetup | SetFitToPages | Aspose.Cells .NET | print worksheet | Excel export | landscape layout
// Common Searches: Aspose.Cells set worksheet orientation landscape C# | FitToPagesWide Aspose.Cells example | How to print Excel sheet on one page width using Aspose.Cells | C# Aspose.Cells PageSetup FitToPagesTall | Aspose.Cells landscape print scaling | SetFitToPages method Aspose.Cells
// Developer Intent: Configure a worksheet’s print settings to use landscape orientation and automatically fit all columns to the page width.
// Use Cases: Print financial statements where a wide column set must appear on a single page without manual scaling. | Generate invoices that require a landscape layout and automatic width fitting for consistent paper output. | Export dashboard reports to Excel for printing on legal‑size paper, ensuring all data fits horizontally.
// AI Prompts: Provide C# code to apply landscape orientation and fit‑to‑width printing to every worksheet in a workbook using Aspose.Cells. | Explain the impact of setting FitToPagesTall to 0 versus a specific number when printing in landscape mode. | Show how to switch between SetFitToPages and individual FitToPagesWide/FitToPagesTall properties in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupDemo
{
    // Shows how to use Aspose.Cells in C# to change a worksheet’s page orientation to landscape and configure PageSetup so all columns fit on one printed page width (FitToPagesWide = 1, FitToPagesTall = 0). The workbook is saved as LandscapeFitToWidth.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set page orientation to Landscape
            worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Enable fit‑to‑width printing:
            // Fit all columns on one page (wide = 1) and let the height adjust automatically (tall = 0)
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.FitToPagesTall = 0;

            // Optionally, you could use the SetFitToPages method instead:
            // worksheet.PageSetup.SetFitToPages(1, 0);

            // Save the workbook to a file
            workbook.Save("LandscapeFitToWidth.xlsx", SaveFormat.Xlsx);
        }
    }
}
