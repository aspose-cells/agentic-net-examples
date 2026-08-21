// Title: Aspose.Cells C# – Export to HTML without hidden worksheets and with FitToCell overflow control
// Description: This example creates a workbook with a visible sheet that contains long wrapped text and a hidden sheet. Using HtmlSaveOptions it sets ExportHiddenWorksheet = false and HtmlCrossStringType = FitToCell, so hidden worksheets are omitted and cell content is confined to the cell width when saved as HTML.
// Keywords: Aspose.Cells HTML export | ExportHiddenWorksheet false | HtmlCrossStringType FitToCell | prevent text overflow HTML | hide hidden worksheets C# | HtmlSaveOptions example | Aspose.Cells C# tutorial
// Common Searches: Aspose.Cells hide hidden sheets when exporting to HTML | FitToCell option to stop text overflow in HTML export | ExportHiddenWorksheet false Aspose.Cells C# | How to limit cell overflow in HTML output with Aspose | Combine ExportHiddenWorksheet and FitToCell in C#
// Developer Intent: Export an Excel workbook to HTML while excluding hidden worksheets and ensuring that long text does not overflow its cell.
// Use Cases: Generate a web‑ready HTML preview of a report where confidential hidden sheets must stay hidden. | Create HTML versions of spreadsheets with fixed column widths, preserving layout consistency across browsers. | Produce printable HTML output where cell content is trimmed to the cell size and hidden data is not exposed.
// AI Prompts: Show how to also remove gridlines in the HTML export while keeping ExportHiddenWorksheet false and FitToCell. | Provide a code snippet that customizes the generated HTML CSS without affecting FitToCell overflow handling. | Explain how to programmatically verify that hidden worksheets are absent from the saved HTML file.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // This example creates a workbook with a visible sheet that contains long wrapped text and a hidden sheet. Using HtmlSaveOptions it sets ExportHiddenWorksheet = false and HtmlCrossStringType = FitToCell, so hidden worksheets are omitted and cell content is confined to the cell width when saved as HTML.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Visible worksheet with data that will overflow
            // -------------------------------------------------
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";

            // Put a long text in A1 that exceeds the column width
            visibleSheet.Cells["A1"].PutValue("This is a very long text that should overflow the cell when exported to HTML.");

            // Apply text wrapping style to demonstrate overflow handling
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;
            StyleFlag flag = new StyleFlag { WrapText = true };
            visibleSheet.Cells["A1"].SetStyle(wrapStyle, flag);
            visibleSheet.Cells.SetColumnWidth(0, 15); // narrow column to force overflow

            // -------------------------------------------------
            // Hidden worksheet (will not be exported)
            // -------------------------------------------------
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");
            hiddenSheet.IsVisible = false; // mark as hidden

            // -------------------------------------------------
            // Configure HTML save options
            // -------------------------------------------------
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Do not export hidden worksheets
                ExportHiddenWorksheet = false,

                // Limit text to the cell width (no overflow)
                HtmlCrossStringType = HtmlCrossType.FitToCell
            };

            // -------------------------------------------------
            // Save the workbook to HTML
            // -------------------------------------------------
            string outputPath = "ExportHiddenAndFitToCell.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to '{outputPath}' with ExportHiddenWorksheet=false and HtmlCrossStringType=FitToCell.");
        }
    }
}
