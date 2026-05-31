using System;
using Aspose.Cells;

namespace ExportExcelToHtmlWithComments
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to cells
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Add a comment to cell B2
            int commentIndex = sheet.Comments.Add("B2");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Discounted price for today";

            // Configure HTML save options to export comments as tooltips
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true,                         // Enable comment export
                ExportCommentsType = PrintCommentsType.PrintInPlace // Render comments as tooltip (in‑place)
            };

            // Save the workbook as an HTML file with comments preserved
            workbook.Save("Products.html", htmlOptions);
        }
    }
}