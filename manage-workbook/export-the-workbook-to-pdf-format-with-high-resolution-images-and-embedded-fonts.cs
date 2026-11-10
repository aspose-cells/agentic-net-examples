// Title: Convert an Excel workbook to PDF while preserving sheet layout and embedding fonts with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF with embedded fonts | how to keep original sheet layout when saving workbook as PDF using Aspose.Cells .NET | C# example for PdfSaveOptions OnePagePerSheet false Aspose.Cells | save Excel workbook to PDF without scaling sheets Aspose.Cells C# | default image resolution for PDF export in Aspose.Cells C#
// Tags: fonts embedding PDF Aspose.Cells | sheet layout preservation PDF Aspose.Cells | OnePagePerSheet false PdfSaveOptions | Excel to PDF conversion C# Aspose.Cells | PdfSaveOptions configuration Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // The sample loads an existing Excel file into an Aspose.Cells Workbook, configures PdfSaveOptions with OnePagePerSheet set to false to retain the original sheet layout, and saves the workbook as a PDF that automatically embeds the required fonts. It includes basic file existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.pdf";

                // Verify that the source workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Source file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Preserve original sheet layout
                    OnePagePerSheet = false
                    // Note: ImageResolution property is not available in this version of Aspose.Cells.
                };

                // Save the workbook as PDF
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
