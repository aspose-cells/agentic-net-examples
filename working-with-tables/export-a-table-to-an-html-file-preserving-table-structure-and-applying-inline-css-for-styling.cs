using System;
using System.Drawing;
using Aspose.Cells;

namespace ExportTableToHtml
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample table data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(5);
            sheet.Cells["C3"].PutValue(0.3);

            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue(0.2);

            // Apply some styling to the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = Color.White;
            headerStyle.ForegroundColor = Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            headerStyle.Borders[BorderType.BottomBorder].Color = Color.Black;

            // Apply the style to the header cells (A1:C1)
            for (int col = 0; col < 3; col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            // Apply border style to the data range
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            for (int row = 1; row <= 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].SetStyle(dataStyle);
                }
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // Export only the table part (no extra worksheet UI)
            htmlOptions.ExportDataOptions = HtmlExportDataOptions.Table;
            // Use inline CSS styles instead of external CSS
            htmlOptions.DisableCss = true;
            // Do not export row/column headings
            htmlOptions.ExportRowColumnHeadings = false;

            // Save the workbook as an HTML file with inline styling
            string outputPath = "TableExport.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}