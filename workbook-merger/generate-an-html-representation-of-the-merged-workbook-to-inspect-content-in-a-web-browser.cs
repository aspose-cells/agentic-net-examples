using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Content of Sheet 1");

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["B2"].PutValue("Content of Sheet 2");

            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
            sheet3.Cells["C3"].PutValue("Content of Sheet 3");

            // Configure HTML save options to export all sheets in a single HTML file
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ShowAllSheets = true,          // Show every sheet
                SaveAsSingleFile = true,       // One HTML file only
                ExportActiveWorksheetOnly = false,
                ExportWorkbookProperties = true,
                ExportWorksheetProperties = true,
                ExportImagesAsBase64 = true    // Embed images directly (if any)
            };

            // Define output path
            string htmlPath = "MergedWorkbook.html";

            // Save the workbook as HTML
            workbook.Save(htmlPath, htmlOptions);

            // Open the generated HTML file in the default web browser for inspection
            Process.Start(new ProcessStartInfo
            {
                FileName = htmlPath,
                UseShellExecute = true
            });
        }
    }
}