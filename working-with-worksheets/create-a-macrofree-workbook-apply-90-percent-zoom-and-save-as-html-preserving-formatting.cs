// Title: Generate a macro‑free workbook, set the first worksheet zoom to 90 %, and export it to HTML while preserving formatting and gridlines using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to create a new workbook without macros, set the zoom of the first worksheet to 90 %, apply bold blue styling to a cell, and save the workbook as an HTML file with gridlines and embedded images. | Show how to configure HtmlSaveOptions in Aspose.Cells for .NET to export all worksheets, embed images as Base64, and retain gridlines when converting a workbook to HTML.
// Common Searches: how to set worksheet zoom level with Aspose.Cells in C# | export Aspose.Cells workbook to HTML with formatting preserved | save workbook as HTML embedding images base64 using Aspose.Cells .NET | create macro‑free Excel file with Aspose.Cells and export to HTML | preserve grid lines when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells set worksheet zoom C# | HtmlSaveOptions embed images base64 Aspose.Cells | export workbook to HTML with gridlines Aspose.Cells | macro‑free workbook creation Aspose.Cells .NET | preserve cell formatting HTML export Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// Creates a new macro‑free Workbook, sets the first worksheet zoom to 90 %, applies bold blue font to cell A1, configures HtmlSaveOptions to embed images as Base64 and retain gridlines, and saves the workbook as output.html.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new macro‑free workbook
            Workbook workbook = new Workbook();

            // Set the zoom level of the first worksheet to 90%
            workbook.Worksheets[0].Zoom = 90;

            // Add sample data with formatting to demonstrate preservation
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            Style style = workbook.CreateStyle();
            style.Font.IsBold = true;
            style.Font.Color = Color.Blue;
            sheet.Cells["A1"].SetStyle(style);

            // Configure HTML save options to preserve formatting
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false, // keep all worksheets
                ExportImagesAsBase64 = true,        // embed images directly
                ExportGridLines = true              // retain grid lines
                // ExportChartImageFormat is not available in current API version
            };

            // Save the workbook as HTML
            workbook.Save("output.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
