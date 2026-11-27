// Title: Create a PDF report showing each worksheet’s original and updated paper size using Aspose.Cells for .NET (C#)
// Common Searches: how to export worksheet paper size changes to a PDF with Aspose.Cells in C# | Aspose.Cells C# generate PDF table of before and after page setup paper size | C# code to list Excel sheet paper size, modify to A4, and save report as PDF using Aspose.Cells | Aspose.Cells save workbook as PDF with custom worksheet information table | retrieve and compare worksheet PaperSizeType values in Aspose.Cells .NET
// Tags: Aspose.Cells retrieve worksheet paper size | Aspose.Cells modify page setup paper size | Aspose.Cells export worksheet data to PDF | Aspose.Cells create custom PDF report from Excel | Aspose.Cells compare before after page setup values

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, records each worksheet's PaperSize before modification, sets the paper size to A4, records the new value, and stores the before/after data. It then builds a new workbook with a table of worksheet names and their paper sizes, applies bold styling to the header, and saves the table as a PDF report.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string reportPath = "PaperSizeReport.pdf";

        // Load the source workbook safely
        Workbook sourceWorkbook;
        try
        {
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            sourceWorkbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Collect worksheet name and paper size information
        var paperSizeInfo = new List<(string SheetName, string Before, string After)>();

        foreach (Worksheet ws in sourceWorkbook.Worksheets)
        {
            // Paper size before modification
            PaperSizeType beforeSize = ws.PageSetup.PaperSize;

            // Modify the paper size (example: set to A4)
            ws.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Paper size after modification
            PaperSizeType afterSize = ws.PageSetup.PaperSize;

            // Store the information
            paperSizeInfo.Add((ws.Name, beforeSize.ToString(), afterSize.ToString()));
        }

        // Create a new workbook for the PDF report
        Workbook reportWorkbook = new Workbook();
        Worksheet reportSheet = reportWorkbook.Worksheets[0];
        reportSheet.Name = "PaperSizeReport";

        // Write table headers
        reportSheet.Cells["A1"].PutValue("Worksheet");
        reportSheet.Cells["B1"].PutValue("Paper Size Before");
        reportSheet.Cells["C1"].PutValue("Paper Size After");

        // Populate the table with collected data
        int currentRow = 1; // zero‑based index; row 1 is the second row
        foreach (var entry in paperSizeInfo)
        {
            reportSheet.Cells[currentRow, 0].PutValue(entry.SheetName);
            reportSheet.Cells[currentRow, 1].PutValue(entry.Before);
            reportSheet.Cells[currentRow, 2].PutValue(entry.After);
            currentRow++;
        }

        // Apply a simple style to the header row
        Style headerStyle = reportWorkbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        StyleFlag styleFlag = new StyleFlag { FontBold = true };
        reportSheet.Cells.CreateRange(0, 0, 1, 3).ApplyStyle(headerStyle, styleFlag);

        // Save the report as a PDF file safely
        try
        {
            reportWorkbook.Save(reportPath, SaveFormat.Pdf);
            Console.WriteLine($"Report saved successfully to {reportPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save PDF report: {ex.Message}");
        }
    }
}
