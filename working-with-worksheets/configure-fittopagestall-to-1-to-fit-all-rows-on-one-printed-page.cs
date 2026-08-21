// Title: Fit all worksheet rows on a single printed page with Aspose.Cells for .NET (C#)
// Description: Shows how to configure Aspose.Cells PageSetup in C# to set FitToPagesTall = 1 and FitToPagesWide = 0, optionally define a print area, and save the workbook so every row prints on one page.
// Keywords: Aspose.Cells | .NET | C# | FitToPagesTall | FitToPagesWide | page scaling | single page print | worksheet print area | Excel printing | PageSetup
// Common Searches: Aspose.Cells fit all rows on one page | FitToPagesTall 1 C# Aspose.Cells example | Print worksheet on a single page using Aspose.Cells | Set FitToPagesWide to 0 Aspose.Cells .NET | Define print area Aspose.Cells PageSetup
// Developer Intent: Configure the worksheet’s PageSetup to limit height only, forcing all rows onto one printed page while allowing unlimited width.
// Use Cases: Create printable reports where the entire data set must appear on a single vertical page. | Generate invoices or receipts that require all line items on one page regardless of column count. | Export large tables to Excel with height‑only scaling for consistent print layouts.
// AI Prompts: Provide a C# code snippet that sets FitToPagesTall to 1 and FitToPagesWide to 0 with Aspose.Cells, including an optional print area. | Explain how to use Aspose.Cells PageSetup to force all rows onto one printed page while preserving column widths. | Show how to calculate the appropriate FitToPagesTall value dynamically based on the number of rows in a worksheet.

using System;
using Aspose.Cells;

// Shows how to configure Aspose.Cells PageSetup in C# to set FitToPagesTall = 1 and FitToPagesWide = 0, optionally define a print area, and save the workbook so every row prints on one page.
class FitToPagesTallDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to make the sheet span many rows
        for (int i = 0; i < 200; i++)
        {
            worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
        }

        // Get the PageSetup object for the worksheet
        PageSetup pageSetup = worksheet.PageSetup;

        // Define the print area (optional, but helps illustrate the scaling)
        pageSetup.PrintArea = "A1:A200";

        // Fit all rows onto a single printed page (tall)
        pageSetup.FitToPagesTall = 1;

        // Set FitToPagesWide to 0 so that only the height is constrained
        pageSetup.FitToPagesWide = 0;

        // Save the workbook to a file
        workbook.Save("FitToPagesTallDemo.xlsx");
    }
}
