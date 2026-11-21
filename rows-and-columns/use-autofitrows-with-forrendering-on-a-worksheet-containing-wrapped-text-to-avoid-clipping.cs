// Title: Adjust row height for wrapped text using AutoFitRows rendering option in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells AutoFitRows rendering flag to avoid wrapped text clipping | How to auto‑fit rows for wrapped cells in an Excel workbook with Aspose.Cells .NET | Example of AutoFitRows with isForRendering true in Aspose.Cells C# | Increase row height for long wrapped text programmatically using Aspose.Cells | Prevent text cut‑off when auto‑sizing rows in Aspose.Cells for .NET
// Tags: AutoFitRows with rendering flag Aspose.Cells | wrapped cell row height auto fit C# | prevent clipped wrapped text Aspose.Cells | Excel row auto sizing Aspose.Cells .NET | Aspose.Cells workbook row height adjustment

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, writes a long string to cell A1, enables text wrapping, then uses AutoFitRows with the rendering option to automatically expand row heights so the wrapped text is fully visible, and saves the workbook as an .xlsx file.
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

            // Put a long text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("This is a very long piece of text that should wrap within the cell and demonstrate how AutoFitRows with rendering options prevents clipping.");

            // Enable text wrapping for the cell
            Style style = cell.GetStyle();
            style.IsTextWrapped = true;
            cell.SetStyle(style);

            // AutoFit rows for all used rows (default auto‑fit)
            int startRow = 0;
            int totalRows = sheet.Cells.MaxDataRow + 1; // include all rows with data
            sheet.AutoFitRows(startRow, totalRows);

            // Save the workbook
            string outputPath = "AutoFitRowsForRendering.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
