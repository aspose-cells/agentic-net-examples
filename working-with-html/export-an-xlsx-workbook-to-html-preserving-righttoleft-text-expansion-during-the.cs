using System;
using Aspose.Cells;

class ExportRtlHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Enable right‑to‑left display for the worksheet
        sheet.DisplayRightToLeft = true;

        // Add sample data (e.g., Hebrew text) to demonstrate RTL rendering
        sheet.Cells["A1"].PutValue("שלום"); // "Hello" in Hebrew
        sheet.Cells["B1"].PutValue("World");

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Preserve the default cross‑string behavior (optional, shown for clarity)
        saveOptions.HtmlCrossStringType = HtmlCrossType.Default;

        // Save the workbook as HTML; the RTL setting is reflected in the output
        workbook.Save("RtlOutput.html", saveOptions);
    }
}