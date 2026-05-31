using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsHtmlSizeComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with data and apply a variety of styles
            for (int row = 0; row < 50; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    cell.PutValue($"R{row + 1}C{col + 1}");

                    // Create a distinct style for each column to increase the style pool
                    Style style = workbook.CreateStyle();
                    style.Font.Name = col % 2 == 0 ? "Arial" : "Times New Roman";
                    style.Font.Size = 10 + (col % 5);
                    style.Font.IsBold = (row % 3 == 0);
                    style.Font.Color = (col % 3 == 0) ? Color.Red : Color.Blue;
                    cell.SetStyle(style);
                }
            }

            // Define output file paths
            string pathWithExclude = "Html_With_ExcludeUnusedStyles.html";
            string pathWithoutExclude = "Html_Without_ExcludeUnusedStyles.html";

            // Save HTML with ExcludeUnusedStyles = true (default)
            HtmlSaveOptions optionsExclude = new HtmlSaveOptions();
            optionsExclude.ExcludeUnusedStyles = true;
            workbook.Save(pathWithExclude, optionsExclude);

            // Save HTML with ExcludeUnusedStyles = false
            HtmlSaveOptions optionsNoExclude = new HtmlSaveOptions();
            optionsNoExclude.ExcludeUnusedStyles = false;
            workbook.Save(pathWithoutExclude, optionsNoExclude);

            // Get file sizes
            long sizeWithExclude = new FileInfo(pathWithExclude).Length;
            long sizeWithoutExclude = new FileInfo(pathWithoutExclude).Length;

            // Calculate reduction
            long reductionBytes = sizeWithoutExclude - sizeWithExclude;
            double reductionPercent = sizeWithoutExclude == 0 ? 0 :
                (double)reductionBytes / sizeWithoutExclude * 100;

            // Output the results
            Console.WriteLine($"HTML size with ExcludeUnusedStyles=true : {sizeWithExclude} bytes");
            Console.WriteLine($"HTML size with ExcludeUnusedStyles=false: {sizeWithoutExclude} bytes");
            Console.WriteLine($"Size reduction: {reductionBytes} bytes ({reductionPercent:F2}%)");
        }
    }
}