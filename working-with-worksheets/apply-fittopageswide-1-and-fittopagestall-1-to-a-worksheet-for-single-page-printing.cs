// Title: Set a worksheet to print on a single page with FitToPagesWide and FitToPagesTall using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that configures a worksheet's PageSetup to fit the content to one page wide and one page tall with Aspose.Cells. | Show how to apply single‑page print scaling (FitToPagesWide = 1, FitToPagesTall = 1) to the first worksheet and save the workbook.
// Common Searches: asp.net how to fit worksheet to one printed page using Aspose.Cells | C# Aspose.Cells set worksheet print scaling to single page | example of page setup for single page printing in Aspose.Cells .NET | Aspose.Cells FitToPagesWide FitToPagesTall usage in C#
// Tags: Aspose.Cells page setup fit to one page | worksheet print scaling C# Aspose.Cells | single page print layout Aspose.Cells .NET | configure worksheet page setup Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, accesses the first worksheet, sets PageSetup.FitToPagesWide and FitToPagesTall to 1 to force the content onto a single printed page, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Apply single-page printing settings
        sheet.PageSetup.FitToPagesWide = 1; // Fit to 1 page wide
        sheet.PageSetup.FitToPagesTall = 1; // Fit to 1 page tall

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
