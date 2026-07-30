// Title: Export an Excel Table to HTML with Inline CSS using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill a range with product data, style the header row, and save only the table portion as an HTML file with embedded CSS, gridlines, and without row/column headings using Aspose.Cells.
// Keywords: Aspose.Cells HTML export C# | export Excel table to HTML | inline CSS Aspose.Cells | HtmlSaveOptions Table export | C# generate HTML from worksheet | save worksheet as HTML without headings
// Common Searches: How to export a specific range to HTML with Aspose.Cells | C# Aspose.Cells export table only HTML | Inline CSS HTML output from Excel using .NET | Remove row and column labels when saving Excel as HTML
// Developer Intent: Generate an HTML representation of a worksheet range that retains formatting, uses inline CSS, and excludes Excel UI elements.
// Use Cases: Create web‑ready product lists directly from Excel data with styled headers. | Produce email‑compatible HTML tables that do not depend on external style sheets. | Build dashboard widgets that display only the data grid, omitting Excel row/column labels.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet range to HTML using inline CSS and no row/column headings. | Explain each HtmlSaveOptions property used for table‑only HTML export in Aspose.Cells. | Provide a step‑by‑step tutorial for styling a header row and saving the sheet as an HTML file with embedded CSS.

using System;
using Aspose.Cells;

namespace ExportTableToHtml
{
    // Demonstrates how to create a workbook, fill a range with product data, style the header row, and save only the table portion as an HTML file with embedded CSS, gridlines, and without row/column headings using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample table data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["C1"].PutValue("Price");

            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(0.5);

            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(5);
            sheet.Cells["C3"].PutValue(0.3);

            // Apply some styling to the header row
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.IsBold = true;
            headerStyle.Font.Color = System.Drawing.Color.White;
            headerStyle.ForegroundColor = System.Drawing.Color.DarkBlue;
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;
            headerStyle.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
            headerStyle.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
            headerStyle.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;

            // Apply the style to the header cells (A1:C1)
            for (int col = 0; col < 3; col++)
            {
                sheet.Cells[0, col].SetStyle(headerStyle);
            }

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export only the table part (no extra workbook UI)
                ExportDataOptions = HtmlExportDataOptions.Table,
                // Use inline CSS styles instead of external CSS files
                DisableCss = true,
                // Include gridlines for better visual fidelity
                ExportGridLines = true,
                // Do not export row/column headings (A, B, 1, 2, etc.)
                ExportRowColumnHeadings = false
            };

            // Save the workbook as an HTML file with the specified options
            string outputPath = "TableExport.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
