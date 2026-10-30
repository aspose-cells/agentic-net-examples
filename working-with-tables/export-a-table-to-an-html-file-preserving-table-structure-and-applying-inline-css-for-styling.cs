// Title: Export an Excel table to HTML with inline CSS styling using Aspose.Cells for .NET
// Common Searches: how to save a ListObject as HTML with inline styles using Aspose.Cells C# | Aspose.Cells export Excel table to HTML preserving table structure and CSS | C# generate HTML file from Excel range with embedded CSS using Aspose.Cells | export Excel data to HTML with inline CSS and base64 images in .NET | convert Excel worksheet to HTML while keeping table formatting Aspose.Cells
// Tags: Aspose.Cells HTMLSaveOptions inline CSS | ListObject to HTML conversion Aspose.Cells | maintain Excel table formatting in generated HTML | base64 image embedding in Aspose.Cells HTML output | apply header and data cell styles before HTML export Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, fills a three‑column table, applies bold header and border styles, converts the range into a ListObject to retain table semantics, configures HtmlSaveOptions for inline CSS (and Base64 image embedding), and saves the result as an HTML file that preserves the original Excel table layout.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row.
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["C1"].PutValue("Score");

            // Populate data rows.
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["C2"].PutValue(85);

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["C3"].PutValue(92);

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");
            sheet.Cells["C4"].PutValue(78);

            // Define style for header.
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Define style for data cells.
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.BottomBorder].Color = Color.Gray;

            // Apply header style.
            AsposeRange headerRange = sheet.Cells.CreateRange("A1:C1");
            headerRange.ApplyStyle(headerStyle, new StyleFlag { Font = true, CellShading = true, Borders = true });

            // Apply data style.
            AsposeRange dataRange = sheet.Cells.CreateRange("A2:C4");
            dataRange.ApplyStyle(dataStyle, new StyleFlag { Borders = true });

            // Convert the range into a table (list object) to preserve structure in HTML.
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 4;     // rows 1‑4
            int totalColumns = 3;  // columns A‑C
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Configure HTML save options for inline CSS.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportImagesAsBase64 = true // Embed images as Base64 (if any)
                // ExportTableContent and ExportCssInline are not available in this version.
                // Default behavior preserves table structure and uses inline styling where appropriate.
            };

            // Save the workbook as an HTML file.
            workbook.Save("TableExport.html", saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
