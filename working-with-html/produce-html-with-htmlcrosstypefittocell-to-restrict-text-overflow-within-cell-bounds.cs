// Title: C# Aspose.Cells – Export to HTML with HtmlCrossType.FitToCell to confine cell text
// Description: Create a workbook, insert a long string, set a narrow column width, and save as HTML using HtmlSaveOptions.HtmlCrossStringType = FitToCell so the text stays within the cell boundaries.
// Keywords: Aspose.Cells HtmlCrossType FitToCell | C# export Excel to HTML | prevent text overflow HTML | HtmlSaveOptions FitToCell example | Aspose.Cells HTML rendering | cell width HTML export
// Common Searches: Aspose.Cells FitToCell HTML export C# | how to stop text overflow when saving Excel as HTML | HtmlCrossStringType FitToCell usage | set column width and fit text in HTML output Aspose | C# Aspose.Cells HTML export options
// Developer Intent: Generate an HTML file from an Aspose.Cells workbook where each cell’s content is limited to the cell’s width by using HtmlCrossType.FitToCell.
// Use Cases: Web‑based reports with narrow columns that must keep long descriptions inside their cells. | Online spreadsheet viewers that require a clean layout without text spilling into adjacent cells. | HTML invoices or catalogs where product details need to stay within predefined column boundaries.
// AI Prompts: Provide a C# snippet that saves an Aspose.Cells workbook to HTML with text confined to each cell using HtmlCrossType.FitToCell. | Explain the effect of HtmlCrossStringType on HTML rendering and how it interacts with column width settings. | Show how to combine column width adjustments and HtmlSaveOptions to prevent overflow in the generated HTML file.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossFitToCellDemo
{
    // Create a workbook, insert a long string, set a narrow column width, and save as HTML using HtmlSaveOptions.HtmlCrossStringType = FitToCell so the text stays within the cell boundaries.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Put a long text into a cell that would normally overflow
            worksheet.Cells["A1"].PutValue("This is a very long text that should be confined within the cell width when exported to HTML.");

            // Optionally set a narrow column width to demonstrate the effect
            worksheet.Cells.SetColumnWidth(0, 10); // Column A width

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set the cross string type to FitToCell to restrict overflow
            htmlOptions.HtmlCrossStringType = HtmlCrossType.FitToCell;

            // Save the workbook as HTML using the specified options
            workbook.Save("FitToCellOutput.html", htmlOptions);

            Console.WriteLine("HTML file saved with HtmlCrossType.FitToCell.");
        }
    }
}
