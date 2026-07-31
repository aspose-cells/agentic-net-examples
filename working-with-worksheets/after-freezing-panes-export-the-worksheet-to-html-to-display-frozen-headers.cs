// Title: C# – Export Worksheet with Frozen Header Row & Column to a Single HTML File using Aspose.Cells
// Description: Demonstrates how to create a workbook, populate sample data, freeze the first row and column, and export the sheet to a single HTML file with frozen headers, row/column headings, and grid lines using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | FreezePanes | HTML export | SaveAsSingleFile | ExportRowColumnHeadings | grid lines | frozen headers | worksheet to HTML | web spreadsheet preview
// Common Searches: Aspose.Cells freeze panes and export to HTML | C# export worksheet with frozen header row to HTML | Save workbook as single HTML file Aspose.Cells | Preserve frozen rows and columns in HTML output | Enable grid lines in Aspose.Cells HTML export
// Developer Intent: Export a worksheet to HTML while keeping the top row and left column fixed during scrolling.
// Use Cases: Generate an HTML report where the header row and first column remain visible while scrolling. | Provide a web‑based spreadsheet preview with frozen headers for better navigation. | Create a single‑file HTML representation of a worksheet that includes grid lines and row/column headings.
// AI Prompts: Write C# code with Aspose.Cells to freeze the first row and column and export the worksheet to a single HTML file that retains frozen headers and shows grid lines. | Explain the impact of SaveAsSingleFile and ExportRowColumnHeadings on frozen pane rendering in Aspose.Cells HTML output. | Give step‑by‑step instructions to change the freeze point and customize HTML export options (e.g., hide grid lines, adjust headings) in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneHtmlExport
{
    // Demonstrates how to create a workbook, populate sample data, freeze the first row and column, and export the sheet to a single HTML file with frozen headers, row/column headings, and grid lines using Aspose.Cells for .NET.
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

            // Freeze the first row and first column (A1 is the freeze point)
            // This will keep the header row and column visible when scrolling in HTML
            sheet.FreezePanes("B2", 1, 1);   // Freeze up to row 1 and column 1

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Save as a single HTML file (required for proper pane handling)
                SaveAsSingleFile = true,
                // Export row and column headings so that frozen headers are rendered
                ExportRowColumnHeadings = true,
                // Optional: show grid lines for better visual reference
                ExportGridLines = true
            };

            // Export the workbook to HTML
            string outputPath = "FrozenPaneExport.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with frozen panes: {outputPath}");
        }
    }
}
