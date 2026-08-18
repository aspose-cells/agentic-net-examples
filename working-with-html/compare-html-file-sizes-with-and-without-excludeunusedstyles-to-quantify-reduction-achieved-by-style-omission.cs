// Title: C# – Compare Aspose.Cells HTML Export Size With and Without ExcludeUnusedStyles
// Description: Creates a workbook with diverse cell styles, saves two HTML files (one with HtmlSaveOptions.ExcludeUnusedStyles = true, the other = false), reads their file sizes, computes the byte and percentage reduction, and writes the results to the console.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExcludeUnusedStyles | .NET | C# | HTML export size | size reduction | style optimization | performance benchmark | spreadsheet to HTML
// Common Searches: Aspose.Cells ExcludeUnusedStyles impact on HTML size | C# compare HTML output size with and without unused style exclusion | measure HTML file size reduction using Aspose.Cells | how much does ExcludeUnusedStyles shrink HTML export | Aspose.Cells HTML size optimization example
// Developer Intent: Find out how enabling ExcludeUnusedStyles changes the size of HTML generated from a workbook and quantify the reduction.
// Use Cases: Benchmark HTML export size for workbooks that contain many unique styles. | Decide whether to enable ExcludeUnusedStyles to lower payload for web delivery. | Validate that style omission does not alter the visual appearance of the exported HTML.
// AI Prompts: Generate a C# snippet that logs HTML file sizes and percentage reduction when using HtmlSaveOptions.ExcludeUnusedStyles. | Explain how to extend the example to also report the size of the accompanying CSS files. | Provide a script that processes multiple workbooks, compares HTML sizes with and without ExcludeUnusedStyles, and outputs a summary CSV.

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsHtmlSizeComparison
{
    // Creates a workbook with diverse cell styles, saves two HTML files (one with HtmlSaveOptions.ExcludeUnusedStyles = true, the other = false), reads their file sizes, computes the byte and percentage reduction, and writes the results to the console.
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
                    style.Font.Size = 10 + (row % 5);
                    style.Font.IsBold = (col % 3 == 0);
                    style.Font.Color = (row % 2 == 0) ? Color.Blue : Color.Green;
                    cell.SetStyle(style);
                }
            }

            // Define output file names
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
