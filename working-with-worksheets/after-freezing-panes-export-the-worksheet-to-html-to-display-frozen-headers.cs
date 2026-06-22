using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneHtmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top 2 rows and the leftmost 2 columns (cell C3 is the freeze point)
            sheet.FreezePanes("C3", 2, 2);

            // Configure HTML save options:
            // - SaveAsSingleFile = true enables proper handling of frozen panes.
            // - ExportRowColumnHeadings = true ensures that the frozen row/column headers are included in the HTML.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                SaveAsSingleFile = true,
                ExportRowColumnHeadings = true
            };

            // Export the worksheet to HTML
            string outputPath = "FrozenPaneExport.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with frozen headers at: {outputPath}");
        }
    }
}