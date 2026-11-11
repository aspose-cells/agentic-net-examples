// Title: How to set A4 printer paper size in an Excel workbook and save it with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet paper size to A4 before saving | Save Excel file with specific printer paper size using Aspose.Cells .NET | Configure PageSetup.PaperSize programmatically in Aspose.Cells workbook | Ensure printed layout matches A4 when exporting Excel with Aspose.Cells | Create workbook, set page setup paper size, and save as XLSX in C#
// Tags: set worksheet paper size Aspose.Cells C# | page setup printer paper size Excel Aspose | save workbook with A4 layout Aspose.Cells | Workbook.Save preserving page setup Aspose | configure print settings Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, sets its PageSetup.PaperSize to PaperA4, ensures the output directory exists, and saves the file as XLSX, preserving the A4 printer paper size for correct print layout.
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

            // Set the printer paper size to A4 (use PaperSizeType for compatibility)
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Define the output file path
            string outputPath = "PrintedWorkbook.xlsx";

            // Determine the directory part of the output path
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                // If no directory is specified, use the current working directory
                outputDir = Directory.GetCurrentDirectory();
                outputPath = Path.Combine(outputDir, outputPath);
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in XLSX format (preserves page setup settings)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
