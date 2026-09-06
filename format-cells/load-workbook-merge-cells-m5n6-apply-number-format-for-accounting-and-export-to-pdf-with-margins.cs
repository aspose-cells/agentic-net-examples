// Title: How to merge cells M5:N6, apply an accounting number format, set PDF margins, and export a worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook, merges the range M5:N6, applies an accounting number format to the merged cells, configures left/right/top/bottom PDF margins, and saves the worksheet as a PDF with Aspose.Cells. | Generate a .NET program that loads or creates a workbook, merges cells M5 through N6, sets a custom accounting style, adjusts page‑setup margins, and exports the sheet to a PDF file using Aspose.Cells.
// Common Searches: Aspose.Cells C# merge cells M5 N6 and export to PDF with custom margins | apply accounting number format to a merged range in Aspose.Cells .NET | set PDF page margins when saving Excel to PDF using Aspose.Cells | how to create a merged cell with accounting style and PDF output in C# | Aspose.Cells save worksheet as PDF with specific left and right margins
// Tags: merge range M5:N6 Aspose.Cells | apply accounting style merged cells Aspose.Cells | configure PDF margins Aspose.Cells | export worksheet to PDF C# Aspose.Cells | custom number format StyleFlag Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// C# example that loads or creates an Excel workbook, merges cells M5:N6, applies a custom accounting number format to the merged range, sets left/right/top/bottom PDF margins, and saves the worksheet as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists; create a blank one if missing
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Save(inputPath); // optional: persist the placeholder file
            }

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells M5:N6 (use fully qualified Aspose.Cells.Range to avoid ambiguity)
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("M5", "N6");
            mergedRange.Merge();

            // Apply accounting number format to the merged range
            Style accountingStyle = mergedRange[0, 0].GetStyle();
            accountingStyle.Custom = "$#,##0.00_);[Red]($#,##0.00)";
            mergedRange.ApplyStyle(accountingStyle, new StyleFlag { NumberFormat = true });

            // Set page margins (in inches)
            sheet.PageSetup.LeftMargin = 0.5;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.7;
            sheet.PageSetup.BottomMargin = 0.7;

            // Export the worksheet to PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
