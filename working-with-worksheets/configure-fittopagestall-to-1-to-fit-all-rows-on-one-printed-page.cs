// Title: Aspose.Cells .NET: FitToPagesTall = 1 to print all rows on a single page (C# example)
// Description: C# sample that creates a workbook, populates 200 rows, defines a print area, sets PageSetup.FitToPagesTall to 1 (height limited to one page) and FitToPagesWide to 0 (auto width), then saves the workbook.
// Keywords: Aspose.Cells | FitToPagesTall | C# page setup | print area | single page print | Excel to PDF | worksheet printing | .NET Excel export
// Common Searches: Aspose.Cells fit all rows on one page | Set FitToPagesTall in C# Aspose.Cells | FitToPagesTall = 1 example | PrintArea and FitToPagesWide usage Aspose.Cells | Aspose.Cells page setup single page height
// Developer Intent: Configure a worksheet so its printed output fits all rows on one page tall.
// Use Cases: Generate a multi‑row report that must fit on a single printed page for PDF generation. | Create an invoice list where rows stay together on one page while column width adjusts automatically. | Prepare a data dump for printing where height is constrained to one page but column count may vary.
// AI Prompts: Show how to set FitToPagesTall to 1 while keeping width unrestricted using Aspose.Cells for .NET. | Provide a C# example that prints a large worksheet on a single page tall, including PrintArea and PDF export. | Explain the impact of setting FitToPagesWide = 0 when FitToPagesTall = 1 in Aspose.Cells page setup.

using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesTallDemo
{
    // C# sample that creates a workbook, populates 200 rows, defines a print area, sets PageSetup.FitToPagesTall to 1 (height limited to one page) and FitToPagesWide to 0 (auto width), then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data that spans many rows
            for (int row = 0; row < 200; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Configure page setup to fit all rows on a single printed page
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.PrintArea = "A1:E200";   // Define the area to be printed
            pageSetup.FitToPagesTall = 1;     // Fit all rows into one page tall
            pageSetup.FitToPagesWide = 0;     // Zero means width is unrestricted (auto‑adjust)

            // Save the workbook to a file
            workbook.Save("FitToPagesTallDemo.xlsx");
        }
    }
}
