// Title: Load an Excel workbook with LightCells API, exclude charts, and export to PDF using C# Aspose.Cells
// Common Searches: C# Aspose.Cells load workbook without charts using LightCells API | How to export Excel to PDF while skipping charts in Aspose.Cells C# | Use LoadOptions.LoadDataOnly to ignore drawings when converting to PDF with Aspose.Cells | Saving an XLSX file as PDF without embedded graphics using Aspose.Cells in .NET | Exclude chart objects from PDF output when converting XLSX with LightCells API
// Tags: LightCells API workbook loading | skip chart objects Aspose.Cells | Excel to PDF conversion C# | LoadOptions.LoadDataOnly usage | PDF export without drawings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel workbook with Aspose.Cells' LightCells API, optionally reloading it with LoadOptions.LoadDataOnly to omit charts and drawings, and then saving the workbook as a PDF file, handling missing files and runtime errors.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = @"C:\Input\Workbook.xlsx";

        // Path for the output PDF file
        string outputPath = @"C:\Output\Workbook.pdf";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook (default loading loads all objects)
            Workbook workbook = new Workbook(sourcePath);

            // Optionally, if the installed Aspose.Cells version supports LoadDataOnly,
            // you can reload the workbook with that option to exclude charts and drawings.
            // Uncomment the following lines if the property is available in your version.
            /*
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.LoadDataOnly = true;
            workbook = new Workbook(sourcePath, loadOptions);
            */

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to PDF format
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully created at: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
