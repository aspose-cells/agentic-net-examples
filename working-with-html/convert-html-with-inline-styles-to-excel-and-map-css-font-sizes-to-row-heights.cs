// Title: Convert an HTML table with inline font-size styles to XLSX and automatically set Excel row heights using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an HTML string containing a table with inline font-size attributes, loads it into an Aspose.Cells Workbook, and then sets each worksheet row’s height to a factor of the largest font size found in that row. | Show how to translate CSS point sizes from HTML cells into appropriate Excel row heights after importing the HTML with Aspose.Cells, and save the workbook as an .xlsx file.
// Common Searches: aspnet convert html table with inline font-size to xlsx using aspose.cells | c# set excel row height based on maximum font size in a row after loading html | how to map css point size to excel row height with Aspose.Cells | adjust worksheet row heights automatically after importing html into workbook in .NET
// Tags: Aspose.Cells HTML to XLSX conversion | map CSS font-size to Excel row height | set worksheet row height from cell style Aspose.Cells | load HTML with inline styles C# Aspose.Cells | automatic row height adjustment based on font size

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads an HTML string that contains a table with inline font-size styles into an Aspose.Cells Workbook, iterates each row to find the maximum font size among its cells, sets the row height to 1.2 times that size, and saves the result as an XLSX file.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Sample HTML with inline styles
            string html = @"
                <html>
                    <body>
                        <table border='1'>
                            <tr>
                                <td style='font-size:12pt;'>Small Text</td>
                                <td style='font-size:16pt;'>Larger Text</td>
                            </tr>
                            <tr>
                                <td style='font-size:10pt;'>Tiny</td>
                                <td style='font-size:14pt;'>Medium</td>
                            </tr>
                        </table>
                    </body>
                </html>";

            // Load HTML into a Workbook using Aspose.Cells
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                // Specify that the source format is HTML
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Get the first worksheet (the HTML is imported here)
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startColumn = usedRange.FirstColumn;
                int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                // Adjust each row height based on the maximum font size in that row
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    double maxFontSize = 0.0;

                    for (int colIndex = startColumn; colIndex <= endColumn; colIndex++)
                    {
                        Cell cell = sheet.Cells[rowIndex, colIndex];
                        if (cell != null && cell.Value != null)
                        {
                            double cellFontSize = cell.GetStyle().Font.Size;
                            if (cellFontSize > maxFontSize)
                            {
                                maxFontSize = cellFontSize;
                            }
                        }
                    }

                    if (maxFontSize > 0)
                    {
                        // Apply a small multiplier for comfortable spacing
                        double rowHeight = maxFontSize * 1.2;
                        sheet.Cells.Rows[rowIndex].Height = rowHeight;
                    }
                }

                // Prepare output path and ensure directory exists
                string outputPath = "ConvertedFromHtml.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                if (string.IsNullOrEmpty(outputDir))
                {
                    // If no directory part, use current directory
                    outputDir = Directory.GetCurrentDirectory();
                    outputPath = Path.Combine(outputDir, outputPath);
                }

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook to an Excel file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
