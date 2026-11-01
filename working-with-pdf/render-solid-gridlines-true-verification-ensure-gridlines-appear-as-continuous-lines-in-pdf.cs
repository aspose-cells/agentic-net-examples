// Title: Render solid gridlines in PDF when converting an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert Excel to PDF with continuous gridlines | Enable solid gridlines in PDF output using PageSetup.PrintGridlines Aspose.Cells | Apply thin black borders to entire used range before saving as PDF in Aspose.Cells | C# code to ensure gridlines appear as solid lines in PDF generated from .xlsx | How to make Excel gridlines print as solid lines in PDF using Aspose.Cells .NET
// Tags: PageSetup.PrintGridlines true Aspose.Cells | cell border style applied with StyleFlag Aspose.Cells | gridline printing enabled for PDF output Aspose.Cells | Workbook.Save PDF with borders Aspose.Cells | C# Excel to PDF conversion using Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX file, enables PageSetup.PrintGridlines, creates a thin black border style applied to the entire used range via a StyleFlag, and saves the workbook as a PDF so that gridlines render as continuous solid lines.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Enable printing of gridlines for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintGridlines = true;

            // Determine the used range
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Create a style with thin black borders on all sides
            Style borderStyle = workbook.CreateStyle();
            borderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            borderStyle.Borders[BorderType.BottomBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.TopBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.LeftBorder].Color = Color.Black;
            borderStyle.Borders[BorderType.RightBorder].Color = Color.Black;

            // Apply only the border elements of the style
            StyleFlag flag = new StyleFlag { Borders = true };

            // Apply the border style to the entire used range
            sheet.Cells.CreateRange(0, 0, maxRow + 1, maxCol + 1).ApplyStyle(borderStyle, flag);

            // Save the workbook as PDF; gridlines will render as solid lines
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during processing or saving: {ex.Message}");
        }
    }
}
