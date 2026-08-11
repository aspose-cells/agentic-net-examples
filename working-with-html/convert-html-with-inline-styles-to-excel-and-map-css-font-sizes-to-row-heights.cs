// Title: C# – Convert HTML with Inline CSS to Excel and Auto‑Adjust Row Height by Font Size (Aspose.Cells)
// Description: The sample reads an HTML document with embedded style attributes, builds a Workbook via HtmlLoadOptions, examines every row to locate the largest font size, multiplies the point size by 1.2 to derive a row height, applies it, and writes an XLSX file. This keeps the original web layout intact in Excel.
// Keywords: Aspose.Cells HTML to Excel | C# inline CSS conversion | map CSS font size to Excel row height | auto row height Aspose.Cells | HtmlLoadOptions example | global developers
// Common Searches: aspocells convert html with inline styles to xlsx | set excel row height based on font size after html import | c# example adjusting row height from css font size | how to preserve html layout in excel using Aspose
// Developer Intent: Transform an HTML document that uses inline CSS into an Excel workbook and automatically set each row’s height to accommodate the biggest font size in that row.
// Use Cases: Create printable Excel reports that retain the visual appearance of web‑based tables. | Migrate email templates or web dashboards with inline styling into spreadsheets without text clipping. | Generate data exports for international teams where row height must reflect varying font sizes.
// AI Prompts: Generate C# code with Aspose.Cells to load an HTML file, compute the maximum font size per row, and set row height proportionally. | Explain the reasoning behind the 1.2 conversion factor from points to row height and propose a more precise formula. | Suggest how to handle merged cells and multi‑line text when calculating row height after importing HTML.

using System;
using Aspose.Cells;

namespace HtmlToExcelConversion
{
    // The sample reads an HTML document with embedded style attributes, builds a Workbook via HtmlLoadOptions, examines every row to locate the largest font size, multiplies the point size by 1.2 to derive a row height, applies it, and writes an XLSX file. This keeps the original web layout intact in Excel.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file that contains inline CSS styles
            string htmlPath = "input.html";

            // Load the HTML file into a workbook.
            // HtmlLoadOptions parses the HTML and creates corresponding cells, styles, etc.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Determine the used range to limit the iteration
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                // Process each row
                for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
                {
                    double maxFontSizeInRow = 0;

                    // Scan all cells in the current row to find the largest font size
                    for (int colIndex = 0; colIndex <= maxCol; colIndex++)
                    {
                        Cell cell = cells[rowIndex, colIndex];
                        if (cell != null && cell.Value != null)
                        {
                            double fontSize = cell.GetStyle().Font.Size;
                            if (fontSize > maxFontSizeInRow)
                                maxFontSizeInRow = fontSize;
                        }
                    }

                    // If any font size was found, adjust the row height accordingly.
                    // The factor 1.2 approximates the conversion from point size to row height.
                    if (maxFontSizeInRow > 0)
                    {
                        Row row = cells.Rows[rowIndex];
                        row.Height = maxFontSizeInRow * 1.2;
                    }
                }
            }

            // Save the workbook as an Excel file.
            string excelPath = "output.xlsx";
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine("HTML has been converted to Excel and row heights adjusted.");
        }
    }
}
