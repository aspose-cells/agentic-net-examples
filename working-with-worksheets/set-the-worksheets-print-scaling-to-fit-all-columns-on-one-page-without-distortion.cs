// Title: How to set worksheet print scaling to fit all columns on one page using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set worksheet to print all columns on a single page | Aspose.Cells PageSetup FitToPagesWide and FitToPagesTall example | fit worksheet columns to page width without scaling rows Aspose.Cells | set worksheet print zoom to 100 percent with Aspose.Cells .NET | configure worksheet print scaling to avoid distortion using Aspose.Cells
// Tags: Aspose.Cells PageSetup FitToPagesWide usage | fit worksheet columns to single page width C# | disable vertical page scaling Aspose.Cells | worksheet print zoom control Aspose.Cells | C# Aspose.Cells print scaling configuration

using Aspose.Cells;
using System;

// The example creates a workbook, accesses the first worksheet, sets PageSetup.FitToPagesWide to 1 so all columns fit the page width, sets FitToPagesTall to 0 to prevent vertical scaling, ensures the zoom is 100%, and saves the file as ScaledWorksheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates a new workbook with a default worksheet

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page setup to fit all columns on one page without distortion
        // FitToPagesWide = 1 forces all columns to fit on a single page width
        // FitToPagesTall = 0 allows rows to span multiple pages (no vertical scaling)
        sheet.PageSetup.FitToPagesWide = 1;
        sheet.PageSetup.FitToPagesTall = 0;

        // Ensure the zoom does not interfere with the fit settings
        sheet.PageSetup.Zoom = 100;

        // Save the workbook to a file
        workbook.Save("ScaledWorksheet.xlsx");
    }
}
