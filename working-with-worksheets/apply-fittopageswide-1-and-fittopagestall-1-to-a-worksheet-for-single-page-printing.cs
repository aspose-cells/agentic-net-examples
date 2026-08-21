// Title: C# Aspose.Cells – Fit Worksheet to One Printed Page (FitToPagesWide = 1, FitToPagesTall = 1)
// Description: This C# example creates a workbook, optionally adds data, and sets Worksheet.PageSetup.FitToPagesWide and FitToPagesTall to 1 so the entire sheet prints on a single page, then saves the file as SinglePagePrint.xlsx.
// Keywords: Aspose.Cells C# FitToPagesWide | FitToPagesTall .NET | single page print Excel | worksheet page scaling | print entire sheet on one page | Aspose.Cells page setup | C# Excel export single page
// Common Searches: Aspose.Cells set FitToPagesWide 1 | Fit worksheet to one page C# | Aspose.Cells single page printing example | How to fit Excel sheet to one page using Aspose.Cells | C# page setup FitToPagesTall
// Developer Intent: Apply FitToPagesWide = 1 and FitToPagesTall = 1 to a worksheet so it prints on a single page.
// Use Cases: Generate printable reports that must fit on one page. | Create invoices or receipts without page breaks. | Export dashboards to Excel with a single‑page layout for easy distribution.
// AI Prompts: Show me how to set FitToPagesWide and FitToPagesTall for all worksheets in a workbook using Aspose.Cells. | Provide code that adjusts page margins, orientation, and scaling together with FitToPagesWide = 1 for single‑page printing. | Explain how to programmatically verify that the page setup will result in a single printed page in Aspose.Cells.

using System;
using Aspose.Cells;

// This C# example creates a workbook, optionally adds data, and sets Worksheet.PageSetup.FitToPagesWide and FitToPagesTall to 1 so the entire sheet prints on a single page, then saves the file as SinglePagePrint.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Add some sample data
        worksheet.Cells["A1"].PutValue("Sample Data 1");
        worksheet.Cells["B1"].PutValue("Sample Data 2");

        // Configure page setup to fit the entire sheet on a single page
        worksheet.PageSetup.FitToPagesWide = 1; // Fit to 1 page wide
        worksheet.PageSetup.FitToPagesTall = 1; // Fit to 1 page tall

        // Save the workbook
        workbook.Save("SinglePagePrint.xlsx");
    }
}
