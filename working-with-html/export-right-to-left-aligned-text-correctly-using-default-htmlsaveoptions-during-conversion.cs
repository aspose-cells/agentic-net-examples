using System;
using Aspose.Cells;

class ExportRtlHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for the worksheet
        worksheet.DisplayRightToLeft = true;

        // Insert right‑to‑left text (Hebrew example) into a cell
        worksheet.Cells["A1"].PutValue("שלום עולם");

        // Ensure the cell's text direction is also set to RightToLeft
        Style rtlStyle = worksheet.Cells["A1"].GetStyle();
        rtlStyle.TextDirection = TextDirectionType.RightToLeft;
        worksheet.Cells["A1"].SetStyle(rtlStyle);

        // Use the default HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as HTML; the RTL alignment will be preserved
        workbook.Save("RtlOutput.html", htmlOptions);
    }
}