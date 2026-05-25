using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook (replace with loading if needed)
        Workbook workbook = new Workbook();

        // Sample worksheets: one with data, two empty
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "DataSheet";
        dataSheet.Cells["A1"].PutValue("Sample Data");

        Worksheet emptySheet1 = workbook.Worksheets.Add("EmptySheet1");
        Worksheet emptySheet2 = workbook.Worksheets.Add("EmptySheet2");

        // Collect indexes of worksheets that contain any data
        List<int> nonEmptyIndexes = new List<int>();
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            // MaxDataRow/MaxDataColumn are -1 when the sheet has no data
            if (ws.Cells.MaxDataRow >= 0 && ws.Cells.MaxDataColumn >= 0)
            {
                nonEmptyIndexes.Add(i);
            }
        }

        // Configure PDF save options to skip empty sheets and avoid blank pages
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Do not output a blank page when a sheet has nothing to print
            OutputBlankPageWhenNothingToPrint = false,
            // Ignore completely blank pages within a sheet
            PrintingPageType = PrintingPageType.IgnoreBlank,
            // Render only the non‑empty worksheets
            SheetSet = new SheetSet(nonEmptyIndexes.ToArray())
        };

        // Save the workbook to PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}