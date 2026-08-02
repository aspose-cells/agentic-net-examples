// Title: Export Multi‑Sheet Workbook to IE‑Compatible MHTML with Aspose.Cells for .NET
// Description: Creates a three‑sheet workbook, configures HtmlSaveOptions (ShowAllSheets, SaveAsSingleFile, IsIECompatible, ExportImagesAsBase64) and saves it as a single MHT file that enables tab switching in Internet Explorer.
// Keywords: Aspose.Cells | .NET | C# | MHTML | IE compatibility | ShowAllSheets | IsIECompatible | ExportImagesAsBase64 | single file export | multi‑sheet HTML | MHT generation
// Common Searches: Aspose.Cells export workbook to MHTML IE compatible | Save multiple worksheets as tabs in a single MHT file | HtmlSaveOptions ShowAllSheets IsIECompatible example | How to embed images as Base64 in MHTML with Aspose.Cells | Internet Explorer tab switching for Aspose.Cells HTML output
// Developer Intent: Produce a single MHTML (MHT) document that contains all workbook sheets as clickable tabs and works correctly in Internet Explorer.
// Use Cases: Legacy web portal needs a self‑contained HTML view of a workbook with sheet tabs for IE users. | Email a multi‑sheet report as an MHT attachment that retains images without external links. | Distribute an offline, single‑file workbook preview that must render correctly in Internet Explorer.
// AI Prompts: Generate C# code to set a custom document title in HtmlSaveOptions when saving as MHTML. | Explain how to programmatically test that the produced MHT file switches sheets correctly in Internet Explorer. | Provide a snippet to reload the saved MHT file into a Workbook and verify that all three worksheets are present.

using System;
using Aspose.Cells;

namespace AsposeCellsMhtmlIE
{
    // Creates a three‑sheet workbook, configures HtmlSaveOptions (ShowAllSheets, SaveAsSingleFile, IsIECompatible, ExportImagesAsBase64) and saves it as a single MHT file that enables tab switching in Internet Explorer.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();

            // First sheet data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "First";
            sheet1.Cells["A1"].PutValue("Content of the first sheet");

            // Second sheet data
            Worksheet sheet2 = workbook.Worksheets.Add("Second");
            sheet2.Cells["B2"].PutValue("Content of the second sheet");

            // Third sheet data
            Worksheet sheet3 = workbook.Worksheets.Add("Third");
            sheet3.Cells["C3"].PutValue("Content of the third sheet");

            // Configure HTML/MHTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export all sheets as tabs
                ShowAllSheets = true,
                // Save as a single file (required for MHTML)
                SaveAsSingleFile = true,
                // Enable IE compatibility mode so that tab switching works in Internet Explorer
                IsIECompatible = true,
                // Export images as Base64 to keep everything inside the MHTML file
                ExportImagesAsBase64 = true
            };

            // Save the workbook as MHTML (MHT) with the above options
            workbook.Save("MultiSheet_IE_Compatible.mht", saveOptions);
        }
    }
}
