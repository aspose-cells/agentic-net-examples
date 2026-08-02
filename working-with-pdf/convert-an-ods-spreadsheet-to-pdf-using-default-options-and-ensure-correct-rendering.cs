// Title: Convert ODS to PDF with Aspose.Cells C# using ConversionUtility (default settings)
// Description: A concise C# example that loads an OpenDocument Spreadsheet (ODS) file and converts it to PDF with Aspose.Cells.Utility.ConversionUtility.Convert, relying on the library's default rendering options for accurate output.
// Keywords: Aspose.Cells ODS to PDF | ConversionUtility default conversion | C# ODS PDF export | OpenDocument spreadsheet PDF .NET | Aspose.Cells PDF rendering | Convert ODS file to PDF C# | Aspose.Cells Utility Convert
// Common Searches: C# convert ODS to PDF Aspose.Cells | Aspose.Cells ConversionUtility example | default ODS to PDF conversion .NET | how to export OpenDocument spreadsheet as PDF using Aspose | Aspose.Cells PDF export without custom settings
// Developer Intent: Generate a PDF from an ODS spreadsheet in C# using Aspose.Cells with default conversion options.
// Use Cases: Batch‑process a folder of ODS files into PDFs on a web server. | Create printable PDF reports from user‑uploaded ODS spreadsheets. | Archive incoming ODS documents as PDFs for compliance and record‑keeping.
// AI Prompts: Show C# code to convert multiple ODS files to PDFs in parallel with Aspose.Cells. | Explain how to add error handling around ConversionUtility.Convert for ODS‑to‑PDF conversion. | Demonstrate customizing page size and margins when converting ODS to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// A concise C# example that loads an OpenDocument Spreadsheet (ODS) file and converts it to PDF with Aspose.Cells.Utility.ConversionUtility.Convert, relying on the library's default rendering options for accurate output.
class Program
{
    static void Main()
    {
        // Path to the source ODS file
        string sourcePath = "input.ods";

        // Desired output PDF file path
        string destPath = "output.pdf";

        // Convert the ODS spreadsheet to PDF using default conversion options.
        // This utilizes Aspose.Cells.Utility.ConversionUtility.Convert(string, string)
        // which handles loading and saving internally.
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("ODS file has been successfully converted to PDF.");
    }
}
