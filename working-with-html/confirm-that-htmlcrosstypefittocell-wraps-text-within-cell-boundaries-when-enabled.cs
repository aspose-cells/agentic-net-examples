// Title: Aspose.Cells .NET: Verify HtmlCrossType.FitToCell wraps text inside cell boundaries on HTML export
// Description: Shows how to enable text wrapping, set HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.FitToCell, and save a workbook as HTML to confirm that long strings remain confined to the cell area.
// Keywords: Aspose.Cells | HtmlCrossType.FitToCell | HTML export | text wrapping | cell overflow prevention | C# | Aspose.Cells example | Workbook to HTML | cell boundary | Aspose.Cells .NET
// Common Searches: HtmlCrossType.FitToCell text wrapping example | Aspose.Cells prevent cell overflow in HTML | How to keep long text inside a cell when exporting to HTML with Aspose.Cells | FitToCell option in Aspose.Cells HTML save | Aspose.Cells HTML export cell width fixed
// Developer Intent: Confirm that setting HtmlCrossStringType to FitToCell forces long cell content to wrap and stay within the cell limits in the generated HTML.
// Use Cases: Creating HTML reports from Excel where column widths must stay constant and text should not spill over. | Generating web‑ready tables from workbooks for dashboards or email templates without layout breakage. | Building printable HTML snapshots of spreadsheets that preserve readability of wrapped cell content.
// AI Prompts: Write a unit test that loads the saved HTML file and asserts that the <td> for A1 contains CSS restricting width and enabling word‑wrap. | Provide a step‑by‑step visual verification guide to ensure the long text does not overflow after using HtmlCrossType.FitToCell. | Explain how to parse the exported HTML with a script to detect style attributes applied by FitToCell (e.g., word-wrap, overflow).

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossTypeDemo
{
    // Shows how to enable text wrapping, set HtmlSaveOptions.HtmlCrossStringType to HtmlCrossType.FitToCell, and save a workbook as HTML to confirm that long strings remain confined to the cell area.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into cell A1 that would normally overflow
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This is a very long text that should be confined within the cell boundaries when HtmlCrossType.FitToCell is used.");

            // Enable text wrapping for the cell (optional, demonstrates interaction with FitToCell)
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Configure HTML save options to use FitToCell cross type
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.HtmlCrossStringType = HtmlCrossType.FitToCell;

            // Save the workbook as HTML
            string outputPath = "FitToCellDemo.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook saved to HTML with HtmlCrossType.FitToCell at: {outputPath}");
        }
    }
}
