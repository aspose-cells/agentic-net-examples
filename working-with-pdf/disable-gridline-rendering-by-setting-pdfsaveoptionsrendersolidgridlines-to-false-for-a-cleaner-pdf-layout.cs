// Title: Hide Excel gridlines in PDF output using Aspose.Cells PdfSaveOptions in C#
// Common Searches: Aspose.Cells C# export Excel to PDF without showing gridlines | Set PdfSaveOptions.RenderSolidGridlines to false example in C# | Remove gridlines from PDF generated from .xlsx using Aspose.Cells | C# code to hide solid gridlines when saving workbook as PDF | How to produce a clean PDF layout from Excel with Aspose.Cells by disabling gridlines
// Tags: Aspose.Cells PdfSaveOptions hide gridlines | Excel workbook to PDF without gridlines C# | RenderSolidGridlines false setting | clean PDF output from Excel Aspose.Cells | disable solid gridlines in PDF export

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, configures PdfSaveOptions with RenderSolidGridlines set to false, and saves the file as a PDF, resulting in a PDF that omits the Excel gridlines for a cleaner layout.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (gridlines are hidden by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
