// Title: Generate a worksheet‑by‑worksheet report of paper size and FitToPages settings with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# list paper size and fit‑to‑page settings for all worksheets in a workbook | how to create a summary sheet of PageSetup properties using Aspose.Cells .NET | retrieve FitToPagesWide and FitToPagesTall values from each Excel sheet with Aspose.Cells | C# code to export worksheet page setup details to a new Excel file using Aspose.Cells | Aspose.Cells generate report of worksheet names, paper size, and fit‑to‑pages options
// Tags: Aspose.Cells worksheet page setup extraction | C# report of Excel paper size per sheet | FitToPagesWide FitToPagesTall retrieval Aspose | generate summary workbook with Aspose.Cells | auto‑fit columns after populating report Aspose

using Aspose.Cells;
using System;

// The example loads an existing workbook, creates a new workbook for the report, writes header cells, then iterates through each worksheet to read PageSetup.PaperSize, FitToPagesWide, and FitToPagesTall. It records the worksheet name and these page‑setup values in the report sheet, auto‑fits the columns, and saves the result as an Excel file.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook sourceWorkbook = new Workbook("input.xlsx");

        // Create a new workbook that will hold the report
        Workbook reportWorkbook = new Workbook();
        Worksheet reportSheet = reportWorkbook.Worksheets[0];
        reportSheet.Name = "Worksheet Report";

        // Write header row
        reportSheet.Cells["A1"].PutValue("Worksheet Name");
        reportSheet.Cells["B1"].PutValue("Paper Size");
        reportSheet.Cells["C1"].PutValue("Fit To Pages Wide");
        reportSheet.Cells["D1"].PutValue("Fit To Pages Tall");

        int rowIndex = 1; // zero‑based index; row 2 in Excel

        // Iterate through each worksheet in the source workbook
        foreach (Worksheet ws in sourceWorkbook.Worksheets)
        {
            PageSetup pageSetup = ws.PageSetup;

            // Retrieve required page‑setup properties
            string paperSize = pageSetup.PaperSize.ToString();          // Paper size (e.g., A4, Letter)
            int fitToPagesWide = pageSetup.FitToPagesWide;              // Number of pages wide
            int fitToPagesTall = pageSetup.FitToPagesTall;              // Number of pages tall

            // Populate the report sheet
            reportSheet.Cells[rowIndex, 0].PutValue(ws.Name);           // Worksheet name
            reportSheet.Cells[rowIndex, 1].PutValue(paperSize);        // Paper size
            reportSheet.Cells[rowIndex, 2].PutValue(fitToPagesWide);   // FitToPagesWide
            reportSheet.Cells[rowIndex, 3].PutValue(fitToPagesTall);   // FitToPagesTall

            rowIndex++;
        }

        // Adjust column widths for better readability
        reportSheet.AutoFitColumns();

        // Save the report workbook (replace with your desired output path)
        reportWorkbook.Save("WorksheetReport.xlsx");
    }
}
