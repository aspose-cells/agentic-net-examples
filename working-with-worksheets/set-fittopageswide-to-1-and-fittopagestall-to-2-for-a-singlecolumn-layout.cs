// Title: Aspose.Cells C# – Set FitToPagesWide = 1 and FitToPagesTall = 2 for a single‑column layout
// Description: C# example that creates a workbook, adds sample data, and uses PageSetup.SetFitToPages to force the worksheet to print on 1 page wide by 2 pages tall, achieving a single‑column layout, then saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# | SetFitToPages | FitToPagesWide | FitToPagesTall | page setup scaling | single column print layout | .NET Excel printing | worksheet fit to pages | Aspose.Cells page scaling example
// Common Searches: Aspose.Cells set FitToPagesWide to 1 | FitToPagesTall = 2 Aspose.Cells .NET | C# page setup fit to 1x2 pages Aspose | how to force single column layout in Excel with Aspose.Cells | Aspose.Cells SetFitToPages example
// Developer Intent: Configure a worksheet’s page setup so the printed output fits exactly one page in width and two pages in height.
// Use Cases: Printing a report that must stay within a single column across two pages for consistent formatting. | Generating invoices where all columns fit on one page width while allowing two pages height for item details. | Creating printable data lists that require a one‑page‑wide, two‑page‑tall layout to preserve column alignment.
// AI Prompts: Show how to set FitToPagesWide and FitToPagesTall separately using PageSetup in Aspose.Cells for .NET. | Provide a C# snippet that scales a worksheet to fit a specific number of pages while keeping the column layout intact. | Explain the differences between PageSetup.SetFitToPages and setting FitToPagesWide/FitToPagesTall directly in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesExample
{
    // C# example that creates a workbook, adds sample data, and uses PageSetup.SetFitToPages to force the worksheet to print on 1 page wide by 2 pages tall, achieving a single‑column layout, then saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some sample data to demonstrate the layout
            worksheet.Cells["A1"].PutValue("Header");
            for (int i = 2; i <= 50; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue($"Row {i - 1}");
            }

            // Set the page to fit 1 page wide and 2 pages tall
            // Using the SetFitToPages method as it directly sets both properties
            worksheet.PageSetup.SetFitToPages(1, 2);

            // Save the workbook to a file
            workbook.Save("FitToPagesExample.xlsx");
        }
    }
}
