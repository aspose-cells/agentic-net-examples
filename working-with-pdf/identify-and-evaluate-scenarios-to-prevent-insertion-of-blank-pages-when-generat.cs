using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing;

namespace AsposeCellsBlankPagePrevention
{
    class Program
    {
        static void Main()
        {
            // Scenario 1: Workbook with a hidden (blank) sheet.
            Workbook wbHiddenSheet = new Workbook();
            Worksheet ws1 = wbHiddenSheet.Worksheets[0];
            ws1.Cells["A1"].PutValue("Visible Sheet Data");
            Worksheet wsHidden = wbHiddenSheet.Worksheets.Add("HiddenSheet");
            wsHidden.IsVisible = false;

            PdfSaveOptions optsHidden = new PdfSaveOptions
            {
                OutputBlankPageWhenNothingToPrint = false
            };
            string pathHidden = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NoBlankPage_HiddenSheet.pdf");
            wbHiddenSheet.Save(pathHidden, optsHidden);

            // Scenario 2: Workbook with sparse data causing blank pages.
            Workbook wbSparse = new Workbook();
            Worksheet wsSparse = wbSparse.Worksheets[0];
            wsSparse.Cells["A1"].PutValue("Header");
            wsSparse.Cells["A2"].PutValue("Row 1");
            wsSparse.Cells["A3"].PutValue("Row 2");
            wsSparse.PageSetup.FitToPagesWide = 1;
            wsSparse.PageSetup.FitToPagesTall = 5;

            PdfSaveOptions optsSparse = new PdfSaveOptions
            {
                PrintingPageType = PrintingPageType.IgnoreBlank,
                OutputBlankPageWhenNothingToPrint = false
            };
            string pathSparse = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NoBlankPage_SparseData.pdf");
            wbSparse.Save(pathSparse, optsSparse);

            // Scenario 3: Workbook with style‑only sheets (no actual cell values).
            Workbook wbStyleOnly = new Workbook();
            Worksheet wsData = wbStyleOnly.Worksheets[0];
            wsData.Cells["A1"].PutValue("Data Cell");
            Worksheet wsStyle = wbStyleOnly.Worksheets.Add("StyleOnly");

            // Apply a background color to a range without putting any values.
            Style style = wbStyleOnly.CreateStyle();
            style.ForegroundColor = Color.LightGray;
            style.Pattern = BackgroundType.Solid;
            wsStyle.Cells.CreateRange("A1:B10").ApplyStyle(style, new StyleFlag() { All = true });

            PdfSaveOptions optsStyle = new PdfSaveOptions
            {
                PrintingPageType = PrintingPageType.IgnoreStyle,
                OutputBlankPageWhenNothingToPrint = false
            };
            string pathStyle = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NoBlankPage_StyleOnly.pdf");
            wbStyleOnly.Save(pathStyle, optsStyle);

            // Scenario 4: Custom page‑saving callback to programmatically skip specific pages.
            Workbook wbCallback = new Workbook();
            Worksheet wsCallback = wbCallback.Worksheets[0];
            for (int i = 0; i < 120; i++)
                wsCallback.Cells[$"A{i + 1}"].PutValue($"Row {i + 1}");

            PdfSaveOptions optsCallback = new PdfSaveOptions
            {
                PageSavingCallback = new SkipEvenPagesCallback(),
                OutputBlankPageWhenNothingToPrint = false
            };
            string pathCallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "NoBlankPage_Callback.pdf");
            wbCallback.Save(pathCallback, optsCallback);

            Console.WriteLine("PDF files generated without unwanted blank pages:");
            Console.WriteLine(pathHidden);
            Console.WriteLine(pathSparse);
            Console.WriteLine(pathStyle);
            Console.WriteLine(pathCallback);
        }

        private class SkipEvenPagesCallback : IPageSavingCallback
        {
            public void PageStartSaving(PageStartSavingArgs args)
            {
                if ((args.PageIndex + 1) % 2 == 0)
                {
                    args.IsToOutput = false;
                }
            }

            public void PageEndSaving(PageEndSavingArgs args)
            {
            }
        }
    }
}