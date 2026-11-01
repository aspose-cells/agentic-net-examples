// Title: Apply AutoFitRows with the ForRendering flag in Aspose.Cells (C#) to prevent wrapped text clipping
// AI Prompts: Generate C# code that creates a workbook, writes a long wrapped string into a cell, enables text wrapping, sets a column width, and calls Worksheet.AutoFitRows with the forRendering parameter to adjust row heights for rendering. | Show how to auto‑fit a specific range of rows using the ForRendering option so that wrapped text remains fully visible when the workbook is rendered to PDF or an image. | Demonstrate extending the example to programmatically determine the maximum data row, apply a custom column width, and save the workbook as an .xlsx file after autofitting rows for rendering.
// Common Searches: Aspose.Cells C# AutoFitRows option to keep wrapped text visible | how to prevent wrapped text clipping when auto‑fitting rows in .NET | adjust row height for long wrapped strings in Excel using Aspose.Cells | C# code sample for AutoFitRows rendering mode to preserve text layout | set column width then auto‑fit rows for rendering in Aspose.Cells workbook
// Tags: auto fit rows with rendering flag Aspose.Cells C# | wrapped text row height adjustment Aspose.Cells | set column width before autofit rows .NET | prevent cell text clipping Aspose.Cells | excel row height auto fit wrapped text C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, writes a long string into cell A1, enables text wrapping, sets column A width to 30 characters, calls AutoFitRows with the ForRendering flag to resize rows based on the wrapped content, and saves the file as AutoFitRows_ForRendering.xlsx.
class AutoFitRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set a long text in cell A1 and enable text wrapping
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This is a very long piece of text that should wrap within the cell and demonstrate how AutoFitRows prevents clipping of wrapped text.");
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // Optionally set column width to a reasonable size
            sheet.Cells.SetColumnWidth(0, 30); // Column A width

            // AutoFit rows (default behavior) to adjust height for wrapped text
            sheet.AutoFitRows(0, sheet.Cells.MaxDataRow + 1);

            // Save the workbook to a file
            workbook.Save("AutoFitRows_ForRendering.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
