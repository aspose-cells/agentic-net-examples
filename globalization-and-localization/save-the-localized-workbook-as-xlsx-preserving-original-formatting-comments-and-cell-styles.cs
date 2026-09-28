// Title: Convert a localized .xls workbook to .xlsx while preserving formatting, comments, and styles using Aspose.Cells for .NET
// AI Prompts: Load a .xls workbook that contains localized content with Aspose.Cells, then save it as .xlsx ensuring all cell formatting, comments, and styles are retained. | Write C# code that checks for the existence of the source Excel file, creates the output directory if missing, and uses Aspose.Cells to export the workbook to XLSX preserving visual elements. | Demonstrate how to use Aspose.Cells SaveFormat.Xlsx to convert an existing workbook while automatically keeping original formatting, comments, and style objects.
// Common Searches: Aspose.Cells keep cell comments when converting localized .xls to .xlsx in C# | C# example to retain sheet formatting while saving workbook as .xlsx with Aspose.Cells | how to export an Excel file to .xlsx without losing localization using Aspose.Cells | create output directory automatically before saving workbook to .xlsx in .NET
// Tags: Aspose.Cells localized workbook to xlsx conversion | maintain visual layout Aspose.Cells | retain workbook comments Aspose.Cells .NET | Xlsx save options Aspose.Cells | ensure destination folder exists C# Excel export

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a localized Excel workbook (.xls), verifies the source file, creates the target folder if needed, and saves the workbook as .xlsx with Aspose.Cells, automatically preserving all original formatting, comments, and cell styles.
class Program
{
    static void Main()
    {
        // Path to the source workbook (any supported format)
        string sourcePath = @"C:\Input\LocalizedWorkbook.xls";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook; all formatting, comments, and styles are retained automatically
            Workbook workbook = new Workbook(sourcePath);

            // Path for the output XLSX file
            string outputPath = @"C:\Output\LocalizedWorkbook_Saved.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in XLSX format (default options preserve formatting, comments, and styles)
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved successfully with original formatting, comments, and styles.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
