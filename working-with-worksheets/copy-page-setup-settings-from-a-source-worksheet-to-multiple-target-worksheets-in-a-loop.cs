// Title: Copy page setup settings from a source worksheet to multiple worksheets in a loop using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy page setup from a source sheet to multiple target sheets | How to duplicate print margins and orientation across worksheets with Aspose.Cells | Loop through worksheets and apply identical page layout settings in Aspose.Cells .NET | Aspose.Cells set same paper size and fit‑to‑page for several sheets programmatically | Copy print area and title rows to other worksheets using Aspose.Cells C#
// Tags: Aspose.Cells copy page setup between worksheets | C# loop set page margins for multiple sheets | duplicate print area Aspose.Cells | apply same paper size to several worksheets C# | transfer page orientation across worksheets Aspose.Cells

using Aspose.Cells;
using System;

// The example loads a workbook, reads the PageSetup of a source worksheet, iterates over a list of target worksheet names, and copies all page‑setup properties—including orientation, paper size, margins, print area, title rows/columns, and fit‑to‑page settings—to each target sheet before saving the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook("SourceWorkbook.xlsx");

        // Source worksheet whose page setup will be copied
        Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // adjust the name as needed

        // Names of target worksheets
        string[] targetSheetNames = { "Sheet2", "Sheet3", "Sheet4" };

        foreach (string sheetName in targetSheetNames)
        {
            Worksheet targetSheet = workbook.Worksheets[sheetName];
            if (targetSheet != null)
            {
                // Copy page setup settings from source to target
                PageSetup src = sourceSheet.PageSetup;
                PageSetup dst = targetSheet.PageSetup;

                dst.Orientation = src.Orientation;
                dst.PaperSize = src.PaperSize;
                dst.TopMargin = src.TopMargin;
                dst.BottomMargin = src.BottomMargin;
                dst.LeftMargin = src.LeftMargin;
                dst.RightMargin = src.RightMargin;
                dst.HeaderMargin = src.HeaderMargin;
                dst.FooterMargin = src.FooterMargin;
                dst.PrintArea = src.PrintArea;
                dst.PrintTitleRows = src.PrintTitleRows;
                dst.PrintTitleColumns = src.PrintTitleColumns;
                dst.FitToPagesWide = src.FitToPagesWide;
                dst.FitToPagesTall = src.FitToPagesTall;
                dst.CenterHorizontally = src.CenterHorizontally;
                dst.CenterVertically = src.CenterVertically;
                // Add additional properties here if more settings are required
            }
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("ResultWorkbook.xlsx");
    }
}
