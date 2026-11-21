// Title: Copy page‑setup settings from the first worksheet to all other worksheets in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy page setup from first sheet to remaining sheets | How to duplicate print margins and orientation for all worksheets with Aspose.Cells | Loop through Excel worksheets to apply same page setup using Aspose.Cells .NET | Synchronize page setup properties across multiple sheets in a workbook with Aspose.Cells
// Tags: Aspose.Cells copy worksheet page setup properties | C# replicate Excel page setup across multiple sheets | Aspose.Cells set paper size and margins for all worksheets | Loop copy print area and orientation with Aspose.Cells | Synchronize page setup settings in Excel workbook using Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook, reads the PageSetup of the first worksheet, then iterates over each remaining worksheet copying all relevant PageSetup properties (paper size, orientation, margins, print area, titles, gridlines, zoom, etc.) before saving the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet's page setup
        PageSetup sourcePageSetup = workbook.Worksheets[0].PageSetup;

        // Loop through all remaining worksheets and copy the page‑setup settings
        for (int i = 1; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];

            // Copy each relevant property from the source page setup
            ws.PageSetup.PaperSize = sourcePageSetup.PaperSize;
            ws.PageSetup.Orientation = sourcePageSetup.Orientation;
            ws.PageSetup.FitToPagesWide = sourcePageSetup.FitToPagesWide;
            ws.PageSetup.FitToPagesTall = sourcePageSetup.FitToPagesTall;
            ws.PageSetup.PrintArea = sourcePageSetup.PrintArea;
            ws.PageSetup.PrintTitleRows = sourcePageSetup.PrintTitleRows;
            ws.PageSetup.PrintTitleColumns = sourcePageSetup.PrintTitleColumns;
            ws.PageSetup.LeftMargin = sourcePageSetup.LeftMargin;
            ws.PageSetup.RightMargin = sourcePageSetup.RightMargin;
            ws.PageSetup.TopMargin = sourcePageSetup.TopMargin;
            ws.PageSetup.BottomMargin = sourcePageSetup.BottomMargin;
            ws.PageSetup.HeaderMargin = sourcePageSetup.HeaderMargin;
            ws.PageSetup.FooterMargin = sourcePageSetup.FooterMargin;
            ws.PageSetup.CenterHorizontally = sourcePageSetup.CenterHorizontally;
            ws.PageSetup.CenterVertically = sourcePageSetup.CenterVertically;
            ws.PageSetup.PrintGridlines = sourcePageSetup.PrintGridlines;
            ws.PageSetup.PrintHeadings = sourcePageSetup.PrintHeadings;
            ws.PageSetup.BlackAndWhite = sourcePageSetup.BlackAndWhite;
            ws.PageSetup.Zoom = sourcePageSetup.Zoom;
            // Add any additional properties you need to copy in the same manner
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
