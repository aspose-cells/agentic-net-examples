using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";

            // Populate some data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Data 2");
            sheet.Cells["B3"].PutValue(456);

            // Merge cells A1:B1 to demonstrate merged cell handling
            sheet.Cells.Merge(0, 0, 1, 2); // Row 0, Column 0, 1 row, 2 columns

            // Apply a style to the merged header cell
            Style headerStyle = sheet.Cells["A1"].GetStyle();
            headerStyle.Font.Name = "Arial";
            headerStyle.Font.Size = 14;
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = System.Drawing.Color.LightBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            sheet.Cells["A1"].SetStyle(headerStyle);

            // Set HTML save options to preserve styles and merged cells
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Export only the active worksheet (the one we just modified)
            saveOptions.ExportActiveWorksheetOnly = true;
            // Keep default behavior for CSS (styles are embedded in the HTML)
            // No need to change ExportWorksheetCSSSeparately unless separate CSS files are desired

            // Define output path (desktop folder for demonstration)
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "WorksheetExport.html");

            // Save the workbook as HTML using the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Worksheet exported to HTML at: {outputPath}");
        }
    }
}