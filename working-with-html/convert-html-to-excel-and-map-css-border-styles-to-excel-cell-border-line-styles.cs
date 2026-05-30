using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsHtmlToExcelDemo
{
    class Program
    {
        static void Main()
        {
            // Paths for the temporary HTML source and the resulting Excel file
            string htmlPath = "sample.html";
            string excelPath = "result.xlsx";

            // Create a simple HTML table with CSS border styles
            string htmlContent = @"
                <html>
                <head>
                    <style>
                        .thinBorder { border: 1px solid #000000; }
                        .mediumBorder { border: 2px solid #FF0000; }
                        .dashedBorder { border: 3px dashed #00FF00; }
                    </style>
                </head>
                <body>
                    <table>
                        <tr>
                            <td class='thinBorder'>Thin</td>
                            <td class='mediumBorder'>Medium</td>
                            <td class='dashedBorder'>Dashed</td>
                        </tr>
                    </table>
                </body>
                </html>";

            // Write the HTML content to a file
            File.WriteAllText(htmlPath, htmlContent);

            // Prepare load options for HTML and save options for XLSX
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            SaveOptions saveOptions = new OoxmlSaveOptions(); // default XLSX save options

            // Convert HTML to Excel using the provided ConversionUtility.Convert method
            ConversionUtility.Convert(htmlPath, loadOptions, excelPath, saveOptions);

            // Load the generated Excel workbook to inspect border styles
            Workbook workbook = new Workbook(excelPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through the first row cells and output their border line styles
            for (int col = 0; col < 3; col++)
            {
                Cell cell = sheet.Cells[0, col];
                Style style = cell.GetStyle();

                Console.WriteLine($"Cell {cell.Name}:");
                Console.WriteLine($"  Top Border    : {style.Borders[BorderType.TopBorder].LineStyle}");
                Console.WriteLine($"  Bottom Border : {style.Borders[BorderType.BottomBorder].LineStyle}");
                Console.WriteLine($"  Left Border   : {style.Borders[BorderType.LeftBorder].LineStyle}");
                Console.WriteLine($"  Right Border  : {style.Borders[BorderType.RightBorder].LineStyle}");
            }

            // Clean up temporary HTML file (optional)
            if (File.Exists(htmlPath))
                File.Delete(htmlPath);
        }
    }
}