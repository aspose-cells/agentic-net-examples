// Title: Configure narrow margins and landscape orientation in Aspose.Cells (C#)
// Description: Demonstrates how to create a new Workbook, set 0.5 cm margins on all sides via the PageSetup object, change the printing orientation to Landscape, and save the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: aspose.cells | c# | page setup | margins | landscape orientation | narrow margins | centimeters | worksheet printing | excel export
// Common Searches: Aspose.Cells set margins in centimeters | C# Aspose.Cells landscape page orientation | how to set narrow margins with Aspose.Cells | page setup margins Aspose.Cells C# example | print Excel workbook landscape using Aspose.Cells
// Developer Intent: Create a workbook with 0.5 cm margins on every side and configure it to print in landscape mode.
// Use Cases: Produce printable reports that maximize data per page by using narrow margins and a landscape layout. | Generate invoices, catalogs, or wide tables where a landscape orientation reduces page breaks. | Export Excel sheets to PDF with fewer pages by applying minimal margins and landscape orientation.
// AI Prompts: Write C# code with Aspose.Cells to set 0.3 cm margins and portrait orientation for a worksheet. | Show how to configure PageSetup margins in inches instead of centimeters using Aspose.Cells. | Explain how to apply the same margin and orientation settings to every worksheet in a workbook programmatically.

using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupDemo
{
    // Demonstrates how to create a new Workbook, set 0.5 cm margins on all sides via the PageSetup object, change the printing orientation to Landscape, and save the file as an XLSX document using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Set narrow margins (values are in centimeters)
            pageSetup.LeftMargin = 0.5;   // 0.5 cm left margin
            pageSetup.RightMargin = 0.5;  // 0.5 cm right margin
            pageSetup.TopMargin = 0.5;    // 0.5 cm top margin
            pageSetup.BottomMargin = 0.5; // 0.5 cm bottom margin

            // Set the page orientation to Landscape
            pageSetup.Orientation = PageOrientationType.Landscape;

            // Save the workbook (uses the provided save rule)
            workbook.Save("NarrowMarginsLandscape.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook created with narrow margins and landscape orientation.");
        }
    }
}
