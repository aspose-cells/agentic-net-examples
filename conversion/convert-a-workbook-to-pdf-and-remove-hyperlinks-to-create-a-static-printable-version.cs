// Title: Convert an Excel workbook to a static PDF and strip all hyperlinks using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx file, deletes every hyperlink, removes hyperlink underline and color styling, and saves the workbook as a PDF. | Demonstrate how to clear the Hyperlinks collection of each worksheet and reset cell styles before calling Workbook.Save with SaveFormat.Pdf.
// Common Searches: aspocells remove hyperlinks before pdf export c# | c# convert excel to pdf without hyperlink formatting using aspose.cells | how to strip hyperlink styling from Excel workbook in .NET | export static printable PDF from Excel workbook Aspose.Cells | clear all hyperlinks in workbook programmatically Aspose.Cells
// Tags: remove hyperlinks Aspose.Cells | excel to pdf conversion without links | clear cell hyperlink formatting Aspose.Cells | static printable PDF from Excel | hyperlink removal before pdf export C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The program loads 'input.xlsx', clears every hyperlink and resets underline and font color on all cells, then saves the workbook as a static 'output.pdf' using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove all hyperlinks from the worksheet
                sheet.Hyperlinks.Clear();

                // Reset hyperlink‑like formatting on cells
                foreach (Cell cell in sheet.Cells)
                {
                    Style style = cell.GetStyle();

                    // Clear underline and set default font color
                    style.Font.Underline = FontUnderlineType.None;
                    style.Font.Color = Color.Black;

                    cell.SetStyle(style);
                }
            }

            // Save the workbook as a PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
