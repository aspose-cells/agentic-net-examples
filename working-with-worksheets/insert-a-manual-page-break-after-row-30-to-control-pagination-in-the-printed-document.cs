// Title: Adding a manual horizontal page break after row 30 in an Excel worksheet with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# add horizontal page break after row 30 | set manual page break in Excel worksheet using Aspose.Cells .NET | control printed page layout in Excel with Aspose.Cells horizontal page breaks | C# insert page break at specific row index using Aspose.Cells | generate PDF from Excel with custom page breaks using Aspose.Cells
// Tags: horizontal page break Aspose.Cells .NET | add page break by row index | Excel worksheet pagination Aspose.Cells | save workbook as PDF with page breaks | manual page break insertion C#

using System;
using Aspose.Cells;

namespace AsposeCellsPageBreakDemo
{
    // The example creates a new workbook, fills the first worksheet with 50 rows of data, inserts a manual horizontal page break at zero‑based row index 30 to split the sheet after row 30, verifies the break, and saves the file both as XLSX and as a PDF to illustrate the pagination effect.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Populate the worksheet with sample data (50 rows)
            Cells cells = worksheet.Cells;
            for (int row = 0; row < 50; row++)
            {
                // Column A: Row number
                cells[row, 0].PutValue($"Row {row + 1}");

                // Column B: Sample numeric value
                cells[row, 1].PutValue(row * 10);
            }

            // Insert a manual horizontal page break after row 30.
            // Aspose.Cells uses zero‑based row indexes for page breaks.
            // Adding a break at index 30 will cause rows 0‑30 to be on the first printed page,
            // and rows 31 onward on the next page.
            worksheet.HorizontalPageBreaks.Add(30);

            // Optional: Verify that the page break was added
            if (worksheet.HorizontalPageBreaks.Count > 0)
            {
                Console.WriteLine($"Horizontal page break added at row index: {worksheet.HorizontalPageBreaks[0].Row}");
            }
            else
            {
                Console.WriteLine("No horizontal page breaks were added.");
            }

            // Save the workbook to an Excel file
            string excelPath = "PageBreakDemo.xlsx";
            workbook.Save(excelPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {excelPath}");

            // Save the workbook to PDF to see the pagination effect (requires Aspose.Cells PDF support)
            string pdfPath = "PageBreakDemo.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved to {pdfPath}");
        }
    }
}
