// Title: C# – Convert ODS Spreadsheet to PDF with Aspose.Cells Default Settings
// Description: Loads an OpenDocument Spreadsheet (ODS) and uses Aspose.Cells.Utility.ConversionUtility to save it as a PDF using the library’s default options, then prints a completion message.
// Keywords: Aspose.Cells | ODS to PDF | ConversionUtility | C# PDF export | OpenDocument Spreadsheet conversion | default conversion settings | Aspose.Cells file format conversion | C# document conversion
// Common Searches: Aspose.Cells convert ODS to PDF C# example | ConversionUtility default PDF export Aspose | How to export OpenDocument Spreadsheet as PDF using Aspose.Cells | C# code for ODS to PDF conversion with Aspose | Save ODS file as PDF with Aspose.Cells utility
// Developer Intent: Transform an ODS file into a PDF using Aspose.Cells without customizing rendering options, within a C# program.
// Use Cases: Batch processing of ODS documents to PDF in a server‑side service. | Generating printable reports from ODS templates without extra configuration. | Automating format conversion in CI/CD pipelines for documentation workflows.
// AI Prompts: Show how to apply custom page margins when converting ODS to PDF with Aspose.Cells. | Add robust error handling for missing or corrupted ODS files during ConversionUtility.Convert. | Demonstrate converting a workbook loaded from a memory stream to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an OpenDocument Spreadsheet (ODS) and uses Aspose.Cells.Utility.ConversionUtility to save it as a PDF using the library’s default options, then prints a completion message.
class Program
{
    static void Main()
    {
        // Path to the source ODS file
        string sourcePath = "input.ods";

        // Desired output PDF file path
        string outputPath = "output.pdf";

        // Convert the ODS spreadsheet to PDF using default options.
        // The ConversionUtility handles loading the ODS file and saving it as PDF.
        ConversionUtility.Convert(sourcePath, outputPath);

        Console.WriteLine("Conversion completed: " + sourcePath + " -> " + outputPath);
    }
}
