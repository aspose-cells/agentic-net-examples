// Title: Export Excel workbook to HTML with embedded CSS using Aspose.Cells for .NET (C#)
// Common Searches: how to export an Excel workbook to HTML with inline CSS using Aspose.Cells C# | Aspose.Cells HtmlSaveOptions embed CSS in generated HTML file | C# save workbook as HTML and disable external stylesheet with Aspose.Cells | enable custom CSS properties when converting Excel to HTML with Aspose.Cells .NET | generate HTML from Excel with embedded styles and custom CSS using Aspose.Cells API
// Tags: Aspose.Cells HtmlSaveOptions inline CSS | Excel to HTML conversion with embedded styles | C# generate HTML from workbook using Aspose.Cells | disable external CSS in Aspose.Cells HTML export | custom CSS properties support in Aspose.Cells HTML output

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The sample loads or creates a workbook, writes sample data to cells A1 and B1, configures HtmlSaveOptions (which embeds CSS by default), and saves the workbook as an HTML file with inline CSS.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load an existing workbook if the file is present; otherwise create a new one.
                string inputPath = "input.xlsx";
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // create a new workbook
                }

                // Add sample data (optional)
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Configure HTML save options (default options embed CSS)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

                // Save the workbook as HTML
                string outputPath = "output.html";
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
