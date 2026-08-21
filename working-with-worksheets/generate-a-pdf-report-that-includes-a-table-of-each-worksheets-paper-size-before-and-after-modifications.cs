// Title: C# – Generate a PDF report of worksheet paper sizes before and after changes using Aspose.Cells for .NET
// Description: Creates a source workbook, records each worksheet’s initial PageSetup.PaperSize, changes all sheets to A4, writes the before/after values into a new workbook, and saves it as a PDF containing a three‑column table (Worksheet, Before Paper Size, After Paper Size).
// Keywords: Aspose.Cells for .NET | C# PDF report | worksheet paper size | PageSetup.PaperSize | before and after paper size | export worksheet layout to PDF | list sheet paper size Aspose.Cells | modify page setup C# | Aspose.Cells SaveFormat.Pdf | automate paper size audit
// Common Searches: Aspose.Cells generate PDF report of sheet paper sizes | C# capture worksheet PageSetup.PaperSize before change | list original and new paper size for each worksheet Aspose.Cells | export before‑after page setup to PDF using Aspose.Cells | how to change all worksheets to A4 and log sizes
// Developer Intent: Produce a PDF that lists each worksheet’s original and updated paper size.
// Use Cases: Audit page‑setup settings across multiple sheets before publishing. | Create a printable summary of layout changes for quality‑control reviews. | Validate that all worksheets conform to a standard paper size in automated pipelines.
// AI Prompts: Write C# code with Aspose.Cells that records each worksheet’s PaperSize, sets all sheets to A4, and saves a PDF containing a table of the before and after sizes. | Explain how to read PageSetup.PaperSize, modify it, and export a summary workbook to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a source workbook, records each worksheet’s initial PageSetup.PaperSize, changes all sheets to A4, writes the before/after values into a new workbook, and saves it as a PDF containing a three‑column table (Worksheet, Before Paper Size, After Paper Size).
class PaperSizeReport
{
    static void Main()
    {
        // Create a source workbook with several worksheets
        Workbook srcWorkbook = new Workbook();
        srcWorkbook.Worksheets.Add("Sheet2");
        srcWorkbook.Worksheets.Add("Sheet3");

        // Set initial paper sizes for demonstration purposes
        srcWorkbook.Worksheets[0].PageSetup.PaperSize = PaperSizeType.PaperLetter;
        srcWorkbook.Worksheets[1].PageSetup.PaperSize = PaperSizeType.PaperA5;
        srcWorkbook.Worksheets[2].PageSetup.PaperSize = PaperSizeType.PaperLegal;

        // Create a new workbook that will hold the PDF report
        Workbook reportWorkbook = new Workbook();
        Worksheet reportSheet = reportWorkbook.Worksheets[0];

        // Write table headers
        reportSheet.Cells["A1"].PutValue("Worksheet");
        reportSheet.Cells["B1"].PutValue("Before Paper Size");
        reportSheet.Cells["C1"].PutValue("After Paper Size");

        // Iterate through each worksheet, capture paper sizes before and after modification
        for (int i = 0; i < srcWorkbook.Worksheets.Count; i++)
        {
            Worksheet ws = srcWorkbook.Worksheets[i];

            // Capture the original paper size
            PaperSizeType beforeSize = ws.PageSetup.PaperSize;

            // Modify the paper size (example: set all to A4)
            ws.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Capture the new paper size
            PaperSizeType afterSize = ws.PageSetup.PaperSize;

            // Populate the report table
            int row = i + 2; // Data starts from row 2
            reportSheet.Cells[row, 0].PutValue(ws.Name);
            reportSheet.Cells[row, 1].PutValue(beforeSize.ToString());
            reportSheet.Cells[row, 2].PutValue(afterSize.ToString());
        }

        // Save the report workbook as a PDF file
        reportWorkbook.Save("PaperSizeReport.pdf", SaveFormat.Pdf);
    }
}
