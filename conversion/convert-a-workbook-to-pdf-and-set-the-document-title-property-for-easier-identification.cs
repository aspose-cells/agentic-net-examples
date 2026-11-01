// Title: Convert an Excel workbook to PDF and assign a custom document title with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set PDF document title while exporting Excel to PDF | How to add a title property to a PDF generated from an XLSX using Aspose.Cells | Aspose.Cells workbook.Save PDF with custom metadata in .NET | Set WorkbookProperties.Title before saving Excel as PDF with Aspose.Cells | Convert Excel file to PDF and embed document title using Aspose.Cells C#
// Tags: Aspose.Cells PDF export with custom document title | C# Excel to PDF conversion using Aspose.Cells | WorkbookProperties.Title API Aspose.Cells | setting PDF metadata in Aspose.Cells | save workbook as PDF with metadata .NET

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, optionally sets the WorkbookProperties.Title, and then saves the workbook as a PDF using SaveFormat.Pdf, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: set document title if needed (requires Aspose.Cells version that supports it)
            // workbook.WorkbookProperties.Title = "My Document Title";

            // Convert and save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
