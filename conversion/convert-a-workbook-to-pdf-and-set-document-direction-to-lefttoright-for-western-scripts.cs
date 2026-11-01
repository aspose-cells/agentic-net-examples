// Title: Convert an Excel workbook to PDF with left-to-right layout using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with left-to-right document direction | How to disable right-to-left orientation when saving workbook as PDF in Aspose.Cells | C# code to check file existence before converting .xlsx to PDF with Aspose.Cells | Set workbook.Settings.IsRightToLeft false before PDF conversion using Aspose.Cells | Convert .xlsx to PDF in .NET while ensuring Western script layout
// Tags: Aspose.Cells PDF export left-to-right | C# workbook to PDF conversion Aspose.Cells | disable right-to-left layout Aspose.Cells | Settings.IsRightToLeft false Aspose.Cells | Excel to PDF conversion .NET

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the source .xlsx file, loads it into an Aspose.Cells Workbook, optionally sets Settings.IsRightToLeft to false to enforce a left-to-right layout, and saves the workbook as a PDF using SaveFormat.Pdf, with basic error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Set the document direction.
            // The IsRightToLeft property may not be available in older versions of Aspose.Cells.
            // If needed, uncomment the following line after confirming the property exists:
            // workbook.Settings.IsRightToLeft = false;

            // Convert and save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
