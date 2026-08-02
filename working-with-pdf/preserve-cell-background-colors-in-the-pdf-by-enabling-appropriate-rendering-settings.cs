using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a value into a cell and set its background color
        var cell = worksheet.Cells["A1"];
        cell.PutValue("Background Color Demo");
        var style = cell.GetStyle();
        style.ForegroundColor = Color.Yellow;          // Desired background color
        style.Pattern = BackgroundType.Solid;          // Apply solid fill
        cell.SetStyle(style);

        // Configure PDF save options – background colors are preserved by default.
        // Setting GridlineColor to Transparent ensures gridlines do not obscure the cell fill.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            GridlineColor = Color.Transparent
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("BackgroundColorDemo.pdf", pdfOptions);
    }
}
// Author: Aspose.Cells .NET example – preserves cell background colors when saving to PDF.