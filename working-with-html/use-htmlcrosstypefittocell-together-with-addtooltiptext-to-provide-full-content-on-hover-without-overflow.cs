// Title: C# Aspose.Cells: Export to HTML with HtmlCrossType.FitToCell and AddTooltipText (hover tooltip)
// Description: Shows how to create a workbook, insert a long string into a cell, narrow the column, and save as HTML using HtmlCrossType.FitToCell so only the visible portion fits the cell while AddTooltipText provides the full text on mouse hover.
// Keywords: Aspose.Cells | HtmlCrossType.FitToCell | AddTooltipText | C# HTML export | cell tooltip | overflow handling | Excel to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells HtmlCrossType.FitToCell tooltip example | C# export Excel to HTML with hover tooltip | AddTooltipText usage in Aspose.Cells | prevent cell overflow in HTML output Aspose.Cells | show full cell text on hover Aspose.Cells HTML
// Developer Intent: Create an HTML file where long cell values are truncated to the column width but reveal the complete content when the user hovers over the cell.
// Use Cases: Web‑based reports with narrow columns that keep full data accessible via tooltips. | Product catalogs or comment sections rendered as HTML tables without breaking layout. | Accessible display of truncated numeric or textual values in spreadsheet‑to‑HTML conversions.
// AI Prompts: Generate C# code to apply a custom CSS class to the tooltip created by AddTooltipText when using HtmlCrossType.FitToCell. | Explain how to handle merged cells while combining HtmlCrossType.FitToCell with AddTooltipText in Aspose.Cells. | Show how to disable the default tooltip and implement a JavaScript‑based hover popup for long cell values.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlFitToCellTooltipDemo
{
    // Shows how to create a workbook, insert a long string into a cell, narrow the column, and save as HTML using HtmlCrossType.FitToCell so only the visible portion fits the cell while AddTooltipText provides the full text on mouse hover.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell that will exceed the column width
            worksheet.Cells["A1"].PutValue("This is a very long text that will not fit into the cell width and we want to see the full content on hover.");

            // Set a narrow column width to force overflow
            worksheet.Cells.SetColumnWidth(0, 10); // Column A width

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Use FitToCell to display only the part that fits inside the cell
            saveOptions.HtmlCrossStringType = HtmlCrossType.FitToCell;
            // Enable tooltip so the full text appears when hovering over the cell
            saveOptions.AddTooltipText = true;

            // Save the workbook as HTML with the specified options
            string outputPath = "FitToCellWithTooltip.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("HTML file saved with HtmlCrossType.FitToCell and tooltip enabled: " + outputPath);
        }
    }
}
