// Title: Render continuous solid gridlines in a PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds thin borders to every cell in the used range and saves the worksheet as a PDF with uninterrupted lines. | Show how to set the PrintGridlines flag and style the full data area to mimic continuous gridlines before exporting to PDF. | Create a workbook, populate sample data, apply a border style across the sheet, enable gridline printing, and generate a PDF file.
// Common Searches: Aspose.Cells C# export Excel to PDF with uninterrupted gridlines | how to make Excel gridlines appear as continuous lines in PDF using Aspose.Cells | enable printing of gridlines and simulate solid borders in PDF conversion Aspose.Cells | apply border style to entire worksheet before PDF export Aspose.Cells .NET
// Tags: continuous line simulation via cell borders | PDF export with uninterrupted gridlines Aspose.Cells | style full worksheet range with borders C# | gridline rendering control in PDF conversion | Aspose.Cells worksheet visual fidelity in PDF

using Aspose.Cells;
using System;

// The example creates a workbook, fills cells with sample data, defines a thin border style and applies it to the entire used range to mimic solid gridlines, sets PageSetup.PrintGridlines to true, and saves the sheet as a PDF named 'SolidGridlines.pdf'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data into cells
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a style with thin solid borders on all sides
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            // Apply the border style to the used range to simulate solid gridlines
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            usedRange.ApplyStyle(borderStyle, new StyleFlag { Borders = true });

            // Ensure that gridlines are printed when exporting to PDF
            sheet.PageSetup.PrintGridlines = true;

            // Save the workbook as a PDF file
            workbook.Save("SolidGridlines.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
