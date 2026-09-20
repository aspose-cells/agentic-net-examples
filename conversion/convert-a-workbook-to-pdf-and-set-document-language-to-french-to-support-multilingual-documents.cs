// Title: Convert an Excel workbook to PDF with French language settings using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, assigns the workbook's CultureInfo to fr-FR, and saves it as a PDF. | Show how to apply a French locale to an Aspose.Cells workbook before exporting to PDF in a .NET application, including basic error handling.
// Common Searches: how to set french locale when exporting Excel to PDF with Aspose.Cells C# | Aspose.Cells convert xlsx to pdf with specific document language | C# example for workbook.Settings.CultureInfo fr-FR before PDF save | multilingual PDF generation from Excel using Aspose.Cells .NET
// Tags: Aspose.Cells set workbook culture fr-FR | Excel to PDF conversion with locale | Workbook.Settings.CultureInfo usage | SaveFormat.Pdf with language setting | C# multilingual PDF export Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;

// The program checks for the presence of input.xlsx, loads it with Aspose.Cells, sets workbook.Settings.CultureInfo to French (fr-FR) for multilingual support, and saves the workbook as output.pdf in PDF format, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the document language to French (fr-FR) for multilingual support
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Convert and save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully converted and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
