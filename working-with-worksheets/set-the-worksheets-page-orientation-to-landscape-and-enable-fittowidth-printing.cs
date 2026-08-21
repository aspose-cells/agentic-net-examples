// Title: C# – Set Worksheet Landscape Orientation and Fit‑to‑Width Printing with Aspose.Cells
// Description: Creates a new Workbook, accesses the first Worksheet, sets PageSetup.Orientation to Landscape, configures FitToPagesWide = 1 (single‑page width) and FitToPagesTall = 0 (auto height), then saves the file as LandscapeFitWidth.xlsx.
// Keywords: Aspose.Cells C# landscape orientation | FitToPagesWide example | FitToPagesTall setting | page setup printing Aspose.Cells | Excel print layout .NET | C# Aspose.Cells worksheet print settings | fit to width Excel | Aspose.Cells page orientation | C# Excel landscape PDF | Aspose.Cells sample code
// Common Searches: Aspose.Cells set landscape orientation C# | fit worksheet to one page width Aspose.Cells | PageSetup properties Aspose.Cells .NET | C# print Excel sheet landscape Aspose | how to use FitToPagesWide in Aspose.Cells
// Developer Intent: Apply landscape orientation and fit‑to‑width printing to a worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a landscape PDF report where all columns fit on a single page width for easy reading. | Print an invoice on legal‑size paper in landscape mode while allowing the height to span multiple pages. | Create a printable schedule that stays horizontally on one page but can extend vertically as needed.
// AI Prompts: Show C# code with Aspose.Cells that sets a worksheet to landscape orientation and fits all columns to one page width. | Explain the PageSetup properties needed for landscape printing and automatic height adjustment in Aspose.Cells. | Provide a step‑by‑step example of combining Landscape orientation with FitToPagesWide = 1 in a .NET workbook.

using System;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, sets PageSetup.Orientation to Landscape, configures FitToPagesWide = 1 (single‑page width) and FitToPagesTall = 0 (auto height), then saves the file as LandscapeFitWidth.xlsx.
class SetPageOrientationAndFitWidth
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set page orientation to Landscape
        worksheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Fit all columns to one page width; let height adjust automatically
        worksheet.PageSetup.FitToPagesWide = 1;   // one page wide
        worksheet.PageSetup.FitToPagesTall = 0;   // auto height

        // Save the workbook
        workbook.Save("LandscapeFitWidth.xlsx");
    }
}
