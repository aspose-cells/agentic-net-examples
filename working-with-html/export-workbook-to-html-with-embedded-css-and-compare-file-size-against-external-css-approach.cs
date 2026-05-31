using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data with formatting
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate cells
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Data 2");
            sheet.Cells["B3"].PutValue(456);

            // Apply some styles to generate CSS
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);

            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Color = System.Drawing.Color.Green;
            sheet.Cells["A2"].SetStyle(dataStyle);
            sheet.Cells["A3"].SetStyle(dataStyle);

            // ---------- Export with embedded CSS (default) ----------
            HtmlSaveOptions embeddedCssOptions = new HtmlSaveOptions();
            // ExportWorksheetCSSSeparately default is false, so CSS will be embedded
            string embeddedHtmlPath = "Workbook_EmbeddedCss.html";
            workbook.Save(embeddedHtmlPath, embeddedCssOptions);

            // ---------- Export with external CSS ----------
            HtmlSaveOptions externalCssOptions = new HtmlSaveOptions();
            externalCssOptions.ExportWorksheetCSSSeparately = true; // CSS will be saved to a separate .css file
            string externalHtmlPath = "Workbook_ExternalCss.html";
            workbook.Save(externalHtmlPath, externalCssOptions);

            // Compare file sizes
            long embeddedSize = new FileInfo(embeddedHtmlPath).Length;
            long externalSize = new FileInfo(externalHtmlPath).Length;

            Console.WriteLine($"Embedded CSS HTML size: {embeddedSize} bytes");
            Console.WriteLine($"External CSS HTML size: {externalSize} bytes");

            // Also display the size of the generated CSS file (if any)
            string cssFilePath = Path.ChangeExtension(externalHtmlPath, ".css");
            if (File.Exists(cssFilePath))
            {
                long cssSize = new FileInfo(cssFilePath).Length;
                Console.WriteLine($"External CSS file size: {cssSize} bytes");
            }
        }
    }
}