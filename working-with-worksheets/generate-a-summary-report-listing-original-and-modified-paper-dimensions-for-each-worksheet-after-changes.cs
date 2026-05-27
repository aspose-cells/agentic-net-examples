using System;
using System.Collections.Generic;
using Aspose.Cells;

class PaperSizeSummary
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add sample worksheets and set initial paper sizes
        Worksheet ws1 = workbook.Worksheets[0];
        ws1.Name = "Sheet1";
        ws1.PageSetup.PaperSize = PaperSizeType.PaperA4; // original size A4

        Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
        ws2.PageSetup.PaperSize = PaperSizeType.PaperLetter; // original size Letter

        // Store original paper dimensions for each worksheet
        List<(string Name, double Width, double Height)> originalInfo = new List<(string, double, double)>();
        foreach (Worksheet ws in workbook.Worksheets)
        {
            originalInfo.Add((ws.Name, ws.PageSetup.PaperWidth, ws.PageSetup.PaperHeight));
        }

        // Modify paper sizes
        ws1.PageSetup.PaperSize = PaperSizeType.PaperA3; // change to A3
        ws2.PageSetup.CustomPaperSize(8.0, 10.0); // custom size 8"x10"

        // Create a summary worksheet
        Worksheet summary = workbook.Worksheets.Add("Summary");

        // Write header row
        summary.Cells[0, 0].PutValue("Worksheet");
        summary.Cells[0, 1].PutValue("Original Width (in)");
        summary.Cells[0, 2].PutValue("Original Height (in)");
        summary.Cells[0, 3].PutValue("Modified Width (in)");
        summary.Cells[0, 4].PutValue("Modified Height (in)");

        // Populate summary data
        for (int i = 0; i < originalInfo.Count; i++)
        {
            var info = originalInfo[i];
            Worksheet ws = workbook.Worksheets[info.Name];
            double modifiedWidth = ws.PageSetup.PaperWidth;
            double modifiedHeight = ws.PageSetup.PaperHeight;

            int row = i + 1;
            summary.Cells[row, 0].PutValue(info.Name);
            summary.Cells[row, 1].PutValue(info.Width);
            summary.Cells[row, 2].PutValue(info.Height);
            summary.Cells[row, 3].PutValue(modifiedWidth);
            summary.Cells[row, 4].PutValue(modifiedHeight);
        }

        // Save the workbook with the summary report
        workbook.Save("PaperSizeSummary.xlsx");
    }
}