using System;
using Aspose.Cells;

namespace AsposeCellsCssCustomPropertiesDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from disk
            // (Replace the path with the actual location of your source file)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to optimize repeated resources (e.g., images)
            htmlOptions.EnableCssCustomProperties = true;

            // Optionally add additional CSS styles (applies when SaveAsSingleFile is true)
            htmlOptions.CssStyles = @"
                body {
                    font-family: Arial, sans-serif;
                    background-color: #f9f9f9;
                    margin: 20px;
                }
                table {
                    border-collapse: collapse;
                }
                td, th {
                    border: 1px solid #ccc;
                    padding: 5px 10px;
                }";

            // Save the workbook as an HTML file using the configured options
            string outputPath = "output.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook converted to HTML with CSS custom properties enabled: {outputPath}");
        }
    }
}