// Title: Aspose.Cells for .NET – Set worksheet to fit 1 page wide and 2 pages tall (single‑column layout)
// Description: Creates a new Workbook, accesses the first Worksheet, configures PageSetup to FitToPagesWide = 1 and FitToPagesTall = 2, turns off percentage scaling, and saves the file as FitToPagesSingleColumn.xlsx.
// Keywords: Aspose.Cells FitToPagesWide | FitToPagesTall .NET example | single column print layout Aspose.Cells | disable percent scaling Aspose.Cells | page setup fit to pages Aspose.Cells
// Common Searches: Aspose.Cells set FitToPagesWide to 1 | FitToPagesTall = 2 Aspose.Cells example | how to disable percent scaling in Aspose.Cells page setup | print Excel worksheet on 1 page wide 2 pages tall using Aspose.Cells
// Developer Intent: Configure a worksheet’s page setup so the printed output fits exactly one page in width and two pages in height without using percentage scaling.
// Use Cases: Generate an Excel file with a predefined single‑column print layout for reports or invoices. | Programmatically enforce a fixed page count for printing when exporting workbooks from a web service. | Adjust existing worksheets to a consistent print format across multiple documents.
// AI Prompts: Write C# code with Aspose.Cells that sets FitToPagesWide = 1, FitToPagesTall = 2, and disables IsPercentScale. | Explain the impact of FitToPagesWide and FitToPagesTall on Excel printing and how to apply them in Aspose.Cells. | Provide a step‑by‑step tutorial for configuring a single‑column page layout and saving the workbook using Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new Workbook, accesses the first Worksheet, configures PageSetup to FitToPagesWide = 1 and FitToPagesTall = 2, turns off percentage scaling, and saves the file as FitToPagesSingleColumn.xlsx.
class SetFitToPagesDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure page setup: fit to 1 page wide and 2 pages tall
        worksheet.PageSetup.FitToPagesWide = 1;
        worksheet.PageSetup.FitToPagesTall = 2;

        // Ensure scaling uses the FitToPages settings
        worksheet.PageSetup.IsPercentScale = false;

        // Save the workbook
        workbook.Save("FitToPagesSingleColumn.xlsx");
    }
}
