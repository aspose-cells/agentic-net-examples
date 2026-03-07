using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class WorksheetVisibilityDemo
{
    static void Main()
    {
        // Create a new workbook with the default sheet
        Workbook workbook = new Workbook();

        // Add two additional worksheets
        workbook.Worksheets.Add("HiddenSheet");
        workbook.Worksheets.Add("VeryHiddenSheet");

        // Hide the second worksheet using the IsVisible property
        workbook.Worksheets[1].IsVisible = false;

        // Hide the third worksheet using the VisibilityType property (VeryHidden)
        workbook.Worksheets[2].VisibilityType = VisibilityType.VeryHidden;

        // Hide the first worksheet using the SetVisible method (ignore errors)
        workbook.Worksheets[0].SetVisible(false, true);

        // Save the workbook normally – hidden states are preserved in the file
        workbook.Save("AllSheets.xlsx");

        // Export only the visible worksheets to HTML
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportHiddenWorksheet = false,   // do not export hidden sheets
            SheetSet = SheetSet.Visible      // render only visible sheets
        };
        workbook.Save("VisibleSheets.html", htmlOptions);

        // Export only the visible worksheets to PPTX
        PptxSaveOptions pptxOptions = new PptxSaveOptions
        {
            SheetSet = SheetSet.Visible      // render only visible sheets
        };
        workbook.Save("VisibleSheets.pptx", pptxOptions);
    }
}