using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Slides;

class ViewModeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];

        // Add some sample data to illustrate different view modes
        ws.Cells["A1"].PutValue("View Mode Demo");
        ws.Cells["A2"].PutValue("Normal View");
        ws.Cells["A3"].PutValue("Page Break Preview");
        ws.Cells["A4"].PutValue("Page Layout View");

        // -------------------------------------------------
        // 1. Set the worksheet to Page Break Preview using ViewType
        // -------------------------------------------------
        ws.ViewType = ViewType.PageBreakPreview;

        // The same effect can be achieved via IsPageBreakPreview property
        ws.IsPageBreakPreview = true;

        // Save the workbook with Page Break Preview view
        workbook.Save("ViewMode_PageBreakPreview.xlsx");

        // -------------------------------------------------
        // 2. Change the view to Page Layout View and enable the ruler
        // -------------------------------------------------
        ws.ViewType = ViewType.PageLayoutView;
        ws.IsRulerVisible = true;

        // Save the workbook with Page Layout view
        workbook.Save("ViewMode_PageLayoutView.xlsx");

        // -------------------------------------------------
        // 3. Export to HTML with different layout modes
        // -------------------------------------------------
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Normal layout (renders like Excel)
        htmlOptions.LayoutMode = HtmlLayoutMode.Normal;
        workbook.Save("ViewMode_Html_Normal.html", htmlOptions);

        // Print layout (renders as if printed)
        htmlOptions.LayoutMode = HtmlLayoutMode.Print;
        workbook.Save("ViewMode_Html_Print.html", htmlOptions);

        // -------------------------------------------------
        // 4. Export to DOCX using AsNormalView (outputs as normal view)
        // -------------------------------------------------
        DocxSaveOptions docxOptions = new DocxSaveOptions
        {
            AsNormalView = true
        };
        workbook.Save("ViewMode_AsNormalView.docx", docxOptions);

        // -------------------------------------------------
        // 5. Export to PPTX with SlideViewType set to Print
        // -------------------------------------------------
        PptxSaveOptions pptxOptions = new PptxSaveOptions
        {
            ExportViewType = SlideViewType.Print,
            OnePagePerSheet = true
        };
        workbook.Save("ViewMode_SlidePrint.pptx", pptxOptions);
    }
}