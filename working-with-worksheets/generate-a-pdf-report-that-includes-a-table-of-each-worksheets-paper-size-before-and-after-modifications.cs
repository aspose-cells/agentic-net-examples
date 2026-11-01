// Title: Generate a PDF report showing each worksheet’s original and updated paper size using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create PDF with table of worksheet paper sizes before and after change | how to log PageSetup PaperSize for each sheet and export to PDF using Aspose.Cells | C# Aspose.Cells generate worksheet paper size comparison report in PDF | save Excel workbook as PDF with custom sheet that lists paper size changes Aspose.Cells | Aspose.Cells .NET capture worksheet PaperSizeType values and include in PDF report
// Tags: Aspose.Cells generate worksheet paper size change summary | C# capture and compare PageSetup PaperSize values | modify worksheet PaperSizeType to A4 with Aspose.Cells | add custom summary sheet and export workbook to PDF | auto-fit columns and apply bold header style in Aspose.Cells sheet

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook with three sheets, records each sheet’s initial PageSetup.PaperSize, changes it to A4, writes the before‑and‑after sizes into a new worksheet named PaperSizeReport, formats the header, auto‑fits columns, and saves the workbook as a PDF file.
class PdfReportGenerator
{
    static void Main()
    {
        try
        {
            // Create a new workbook with sample worksheets
            Workbook workbook = new Workbook();

            // Add three worksheets with different initial paper sizes
            Worksheet ws1 = workbook.Worksheets[0];
            ws1.Name = "Sheet1";
            ws1.PageSetup.PaperSize = PaperSizeType.PaperA5; // Initial size A5

            Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
            ws2.PageSetup.PaperSize = PaperSizeType.PaperLegal; // Initial size Legal

            Worksheet ws3 = workbook.Worksheets.Add("Sheet3");
            ws3.PageSetup.PaperSize = PaperSizeType.PaperLetter; // Initial size Letter

            // List to hold report rows: (Worksheet name, Before size, After size)
            List<(string SheetName, PaperSizeType Before, PaperSizeType After)> reportData =
                new List<(string, PaperSizeType, PaperSizeType)>();

            // Iterate through each worksheet, capture before size, modify, capture after size
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // Capture the paper size before modification
                PaperSizeType beforeSize = ws.PageSetup.PaperSize;

                // Modify the paper size to A4 for demonstration
                ws.PageSetup.PaperSize = PaperSizeType.PaperA4;

                // Capture the paper size after modification
                PaperSizeType afterSize = ws.PageSetup.PaperSize;

                // Store the data for the report
                reportData.Add((ws.Name, beforeSize, afterSize));
            }

            // Create a new worksheet to hold the report table
            Worksheet reportSheet = workbook.Worksheets.Add("PaperSizeReport");

            // Write table headers
            reportSheet.Cells["A1"].PutValue("Worksheet");
            reportSheet.Cells["B1"].PutValue("Paper Size Before");
            reportSheet.Cells["C1"].PutValue("Paper Size After");

            // Populate the table with collected data
            int rowIndex = 1; // zero‑based index; row 1 is the second row in the sheet
            foreach (var entry in reportData)
            {
                reportSheet.Cells[rowIndex, 0].PutValue(entry.SheetName);
                reportSheet.Cells[rowIndex, 1].PutValue(entry.Before.ToString());
                reportSheet.Cells[rowIndex, 2].PutValue(entry.After.ToString());
                rowIndex++;
            }

            // Apply simple formatting (bold header)
            Style headerStyle = reportSheet.Cells["A1:C1"].GetStyle();
            headerStyle.Font.IsBold = true;
            reportSheet.Cells["A1:C1"].SetStyle(headerStyle);

            // Adjust column widths to fit content
            reportSheet.AutoFitColumns();

            // Define output PDF path
            string outputPath = "PaperSizeReport.pdf";

            // Save the workbook as a PDF; the report sheet will be included in the PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF report generated successfully: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
