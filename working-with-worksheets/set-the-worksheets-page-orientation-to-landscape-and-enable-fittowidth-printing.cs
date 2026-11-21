// Title: Set worksheet page orientation to landscape and enable fit‑to‑width printing using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to print a worksheet in landscape mode and fit to one page wide | C# code sample for setting PageSetup Orientation to Landscape with FitToPagesWide = 1 in Aspose.Cells | Configure worksheet printing options to landscape and fit‑to‑width in Aspose.Cells .NET | Aspose.Cells page setup landscape orientation and fit to width example | Set worksheet print area to landscape and auto‑scale width using Aspose.Cells in C#
// Tags: Aspose.Cells set worksheet orientation landscape | Aspose.Cells fit worksheet to one page width | C# PageSetup FitToPagesWide example | Aspose.Cells printing options configuration | Aspose.Cells page orientation landscape setting

using System;
using Aspose.Cells;

// The example creates a workbook, accesses the first worksheet, sets its page orientation to Landscape, configures the PageSetup to fit the content to one page wide with unlimited height, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set page orientation to Landscape
        sheet.PageSetup.Orientation = PageOrientationType.Landscape;

        // Enable Fit‑to‑Width printing (fit to 1 page wide, unlimited height)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0; // 0 means no restriction on height

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
