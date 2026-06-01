using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlExportWithPageBreakComments
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook and populate it with sample data
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Fill some rows and columns with data
                for (int row = 0; row < 100; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // ------------------------------------------------------------
                // 2. Define horizontal page breaks at row 30 and row 60
                // ------------------------------------------------------------
                // In Aspose.Cells the collection is called HorizontalPageBreaks.
                sheet.HorizontalPageBreaks.Add(30);
                sheet.HorizontalPageBreaks.Add(60);

                // ------------------------------------------------------------
                // 3. Configure HTML save options
                //    - SaveAsSingleFile = true  : all worksheets in one HTML file
                //    - ExportPageHeaders = false (not needed for this demo)
                // ------------------------------------------------------------
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    SaveAsSingleFile = true,
                    ExportPageHeaders = false,
                    // Ensure that page breaks are respected during export
                    ExportPrintAreaOnly = false
                };

                // ------------------------------------------------------------
                // 4. Save the workbook to a memory stream (HTML content in memory)
                // ------------------------------------------------------------
                using (MemoryStream htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, saveOptions);

                    // Convert the stream to a string for post‑processing
                    string htmlContent = Encoding.UTF8.GetString(htmlStream.ToArray());

                    // ------------------------------------------------------------
                    // 5. Insert HTML comments as page‑break markers
                    //    Aspose.Cells inserts a <div class="pagebreak"> element before each
                    //    page break when SaveAsSingleFile is true. We replace that element
                    //    with a comment marker: <!--PageBreak-->
                    // ------------------------------------------------------------
                    const string pageBreakDiv = "<div class=\"pagebreak\"></div>";
                    const string commentMarker = "<!--PageBreak-->";

                    htmlContent = htmlContent.Replace(pageBreakDiv, commentMarker);

                    // ------------------------------------------------------------
                    // 6. Write the modified HTML to a file
                    // ------------------------------------------------------------
                    string outputPath = "WorkbookWithPageBreakComments.html";

                    // Ensure the directory exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    File.WriteAllText(outputPath, htmlContent, Encoding.UTF8);

                    Console.WriteLine($"HTML file saved to: {Path.GetFullPath(outputPath)}");
                    Console.WriteLine("Page break markers have been inserted as HTML comments.");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}