// Title: Merge cells M5:N6, apply accounting number format, set page margins, and export worksheet to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge cells M5:N6 and apply accounting format before PDF export | Set custom page margins when saving Excel worksheet as PDF with Aspose.Cells .NET | Apply accounting number style to a merged range in Aspose.Cells and generate PDF | C# code to merge a range, format numbers, and export to PDF using Aspose.Cells | How to use PdfSaveOptions with custom margins in Aspose.Cells for .NET
// Tags: merge cell range M5:N6 Aspose.Cells | accounting number format style Aspose.Cells | custom page margins PDF export Aspose.Cells | PdfSaveOptions margin settings Aspose.Cells | C# Excel to PDF conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range (C# 8.0 feature)
using AsposeRange = Aspose.Cells.Range;

// Loads an existing Excel workbook, merges cells M5:N6 on the first worksheet, applies an accounting number format to the merged range, configures 0.5‑inch page margins, and saves the worksheet as a PDF using PdfSaveOptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create a range for cells M5:N6 and merge them
            AsposeRange mergeRange = sheet.Cells.CreateRange("M5:N6");
            mergeRange.Merge();

            // Apply accounting number format to the merged range
            Style accountingStyle = workbook.CreateStyle();
            accountingStyle.Custom = "$#,##0.00;[Red]($#,##0.00)"; // accounting format

            StyleFlag styleFlag = new StyleFlag
            {
                NumberFormat = true
            };
            mergeRange.ApplyStyle(accountingStyle, styleFlag);

            // Set page margins (in inches)
            sheet.PageSetup.LeftMargin = 0.5;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;

            // Export the worksheet to PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            try
            {
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
