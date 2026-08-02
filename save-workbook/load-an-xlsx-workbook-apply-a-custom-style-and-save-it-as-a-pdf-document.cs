// Title: Apply a Custom Cell Style to an XLSX Workbook and Export as PDF with Aspose.Cells for .NET (C#)
// Description: C# example that loads an XLSX file (creates a minimal one if missing), defines a custom style (Arial 12, yellow solid background), applies it to cells A1:B2, and saves the workbook as a PDF using Aspose.Cells.
// Keywords: Aspose.Cells C# PDF conversion | apply custom style Aspose.Cells | Excel to PDF with styling | create style programmatically Aspose.Cells | range formatting Aspose.Cells | save workbook as PDF .NET | cell background color Aspose.Cells
// Common Searches: How to apply a custom style to a range in Excel using Aspose.Cells C# | Convert styled XLSX to PDF with Aspose.Cells for .NET | Aspose.Cells example: set cell background and export to PDF | C# code to create and apply a style before PDF export | Aspose.Cells PDF export with cell formatting
// Developer Intent: Load an existing XLSX workbook, format a specific cell range with a custom style, and generate a PDF version of the workbook.
// Use Cases: Produce branded PDF reports by highlighting key cells in an Excel template. | Create printable invoices where header rows are emphasized with a custom background before PDF conversion. | Automate archival of spreadsheets with highlighted audit fields, exporting them to PDF for compliance.
// AI Prompts: Show how to apply the custom style to the entire worksheet before saving as PDF. | Demonstrate setting different font colors for multiple ranges and exporting each range to separate PDF files. | Explain how to retain cell borders, alignment, and other formatting when converting a styled workbook to PDF with Aspose.Cells.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// C# example that loads an XLSX file (creates a minimal one if missing), defines a custom style (Arial 12, yellow solid background), applies it to cells A1:B2, and saves the workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Paths for source Excel file and destination PDF file
        string sourcePath = "input.xlsx";
        string destPath   = "output.pdf";

        try
        {
            // Ensure the source file exists; create a minimal workbook if it does not
            if (!File.Exists(sourcePath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Save(sourcePath);
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 12;
            customStyle.ForegroundColor = Color.Yellow;
            customStyle.Pattern = BackgroundType.Solid;

            // Define which style attributes to apply
            StyleFlag styleFlag = new StyleFlag { All = true };

            // Apply the custom style to a range of cells (A1:B2)
            Aspose.Cells.Range range = workbook.Worksheets[0].Cells.CreateRange("A1:B2");
            range.ApplyStyle(customStyle, styleFlag);

            // Save the workbook as PDF
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Workbook loaded, style applied, and saved as PDF successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
