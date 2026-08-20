// Title: Export a worksheet with frozen panes to a single HTML file using Aspose.Cells for .NET
// Description: Demonstrates how to freeze the first row and column, configure HtmlSaveOptions (SaveAsSingleFile, ExportRowColumnHeadings, ExportGridLines), and save the worksheet as an HTML page that retains the frozen headers for web viewing.
// Keywords: Aspose.Cells freeze panes HTML export | C# HtmlSaveOptions SaveAsSingleFile | ExportRowColumnHeadings Aspose.Cells | Aspose.Cells grid lines HTML | freeze first row and column Aspose.Cells | Aspose.Cells worksheet to HTML
// Common Searches: Aspose.Cells export frozen panes to HTML | C# save Excel with frozen headers as single HTML file | HtmlSaveOptions keep frozen rows visible | How to export Excel with frozen columns using Aspose.Cells | Export worksheet with grid lines and frozen headers
// Developer Intent: Create an HTML representation of an Excel worksheet that keeps the top row and left column fixed, mimicking Excel's frozen pane behavior.
// Use Cases: Web dashboards where header rows/columns must stay visible while scrolling large data sets. | Generating printable HTML previews of Excel reports that retain frozen pane layout. | Embedding Excel-like tables in documentation portals without requiring the Excel application.
// AI Prompts: Write C# code with Aspose.Cells to freeze the first row and column and export the sheet to a single HTML file that includes row/column headings and grid lines. | Explain the impact of HtmlSaveOptions properties SaveAsSingleFile, ExportRowColumnHeadings, and ExportGridLines on the HTML output of a frozen‑pane worksheet. | Create a unit test in C# that verifies the exported HTML contains the frozen header rows and columns after using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneHtmlExport
{
    // Demonstrates how to freeze the first row and column, configure HtmlSaveOptions (SaveAsSingleFile, ExportRowColumnHeadings, ExportGridLines), and save the worksheet as an HTML page that retains the frozen headers for web viewing.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (10 rows x 5 columns)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the first row and first column (A2 cell is the freeze point)
            // This will keep row 1 and column A visible while scrolling
            sheet.FreezePanes("B2", 1, 1);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Save as a single HTML file so that page headers can be exported if needed
                SaveAsSingleFile = true,
                // Export row and column headings to make the frozen headers visible in HTML
                ExportRowColumnHeadings = true,
                // Optional: include grid lines for better visual reference
                ExportGridLines = true
            };

            // Export the worksheet to HTML
            string outputPath = "FrozenPaneExport.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with frozen headers at: {outputPath}");
        }
    }
}
