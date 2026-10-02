// Title: Set 0.5‑inch page margins for all worksheets and export an Excel workbook to a single HTML file with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx workbook, applies 0.5‑inch margins to every worksheet, and saves it as one HTML document with all images embedded as Base64 using Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to embed images, include all worksheets, and keep the original margins when converting Excel to HTML in C#.
// Common Searches: how to export an Excel file to HTML with embedded images using Aspose.Cells .NET | C# set uniform page margins for all worksheets before saving as HTML | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example | save multiple worksheets to one HTML document with base64 images in C# | preserve page setup margins when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions embed images base64 | set worksheet page margins Aspose.Cells | export workbook as one HTML page Aspose.Cells | preserve page setup during HTML conversion | C# convert Excel to HTML with embedded images

using System;
using Aspose.Cells;

// Loads input.xlsx, applies 0.5‑inch margins to each worksheet, configures HtmlSaveOptions to embed images as Base64 and export all worksheets, then saves the result as a single HTML file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Enable and configure page margins for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            // Margins are specified in inches
            pageSetup.LeftMargin = 0.5;
            pageSetup.RightMargin = 0.5;
            pageSetup.TopMargin = 0.5;
            pageSetup.BottomMargin = 0.5;
        }

        // Set HTML save options to embed images directly in the HTML as base64 strings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;          // Embed images
        htmlOptions.ExportActiveWorksheetOnly = false;   // Export all worksheets

        // Save the workbook as an HTML file with embedded images
        workbook.Save("output.html", htmlOptions);
    }
}
