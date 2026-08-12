// Title: C# Export Excel to HTML with Aspose.Cells – HtmlCrossType.FitToCell to Prevent Text Overflow
// Description: Demonstrates how to create a workbook, insert a long string into a cell, set a narrow column width, and save the file as HTML using Aspose.Cells HtmlSaveOptions with HtmlCrossType.FitToCell so the text is truncated to the cell's width instead of spilling over.
// Keywords: Aspose.Cells HtmlCrossType FitToCell | C# HTML export Excel | prevent text overflow Aspose.Cells | set column width Excel C# | HtmlSaveOptions FitToCell example
// Common Searches: Aspose.Cells fit text to cell when exporting HTML | HtmlCrossType.FitToCell C# sample | how to stop cell overflow in HTML output Aspose | export Excel to HTML with controlled column width
// Developer Intent: Generate an HTML file from an Excel workbook where cell content is confined to the cell width using HtmlCrossType.FitToCell.
// Use Cases: Web dashboards that require fixed‑width columns to keep layout tidy. | Email‑ready HTML tables where long descriptions must not break the design. | Printable HTML snapshots of spreadsheets that preserve column boundaries.
// AI Prompts: Show how to apply HtmlCrossType.FitToCell to every worksheet in a workbook when saving to HTML with Aspose.Cells. | Compare HtmlCrossType.FitToCell with other HtmlCrossStringType options and suggest ideal scenarios for each. | Provide code that adjusts row height after using FitToCell to avoid clipped text in the generated HTML.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossFitToCellDemo
{
    // Demonstrates how to create a workbook, insert a long string into a cell, set a narrow column width, and save the file as HTML using Aspose.Cells HtmlSaveOptions with HtmlCrossType.FitToCell so the text is truncated to the cell's width instead of spilling over.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell that would normally overflow
            worksheet.Cells["A1"].PutValue("This is a very long text that should be truncated to fit within the cell width.");

            // Set a narrow column width to force overflow
            worksheet.Cells.SetColumnWidth(0, 10); // Column A width

            // Configure HTML save options to use FitToCell cross type
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HtmlCrossStringType = HtmlCrossType.FitToCell;

            // Save the workbook as HTML with the specified options
            workbook.Save("FitToCellOutput.html", htmlOptions);

            Console.WriteLine("HTML file saved with HtmlCrossType.FitToCell.");
        }
    }
}
