using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgToHtmlDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------------------------------------------------
                // 1. Create a workbook and add sample data to the first sheet
                // -------------------------------------------------------------
                Workbook workbook = new Workbook();                     // create a new workbook
                Worksheet sheet = workbook.Worksheets[0];              // get the first worksheet

                // Populate cells
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Quantity");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(85);

                // -------------------------------------------------------------
                // 2. Render the worksheet to an SVG file
                // -------------------------------------------------------------
                SvgImageOptions svgOptions = new SvgImageOptions();    // options for SVG rendering
                svgOptions.FitToViewPort = true;                      // fit SVG to viewport

                // Render the first page of the sheet to SVG
                SheetRender renderer = new SheetRender(sheet, svgOptions);
                string svgFileName = "worksheet.svg";
                renderer.ToImage(0, svgFileName);                     // save SVG file

                // -------------------------------------------------------------
                // 3. Create a simple HTML document that embeds the SVG
                // -------------------------------------------------------------
                string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>Worksheet SVG in HTML</title>
    <style>
        .svg-container {{
            width: 100%;
            height: 80vh;
        }}
    </style>
</head>
<body>
    <h2>Worksheet rendered as SVG</h2>
    <div class=""svg-container"">
        <object type=""image/svg+xml"" data=""{svgFileName}"" width=""100%"" height=""100%"">
            Your browser does not support SVG.
        </object>
    </div>
</body>
</html>";

                // Save the HTML file
                string htmlFileName = "worksheet.html";
                File.WriteAllText(htmlFileName, htmlContent);

                // -------------------------------------------------------------
                // 4. (Optional) Save the original workbook for reference
                // -------------------------------------------------------------
                workbook.Save("original_workbook.xlsx");

                Console.WriteLine($"SVG generated: {Path.GetFullPath(svgFileName)}");
                Console.WriteLine($"HTML generated: {Path.GetFullPath(htmlFileName)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}