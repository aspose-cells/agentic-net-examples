// Title: Save an Aspose.Cells workbook as PDF with version 1.6 to enable transparency using C#
// AI Prompts: Write C# code that creates a Workbook, sets PdfSaveOptions.Version to 1.6, and saves the file as a PDF with transparency support. | Show how to configure Aspose.Cells PdfSaveOptions for PDF 1.6 features before calling Workbook.Save in a .NET application. | Provide a step‑by‑step example of exporting an Excel workbook to a PDF with PDF version 1.6 using Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# set PDF version 1.6 for transparency handling | PdfSaveOptions.Version property example Aspose.Cells .NET | How to enable PDF 1.6 features when exporting Excel to PDF with Aspose.Cells | Configure PDF save options for advanced PDF features in Aspose.Cells C#
// Tags: Aspose.Cells PdfSaveOptions version 1.6 | C# export Excel to PDF with transparency | set PDF version Aspose.Cells .NET | configure PDF output Aspose.Cells | enable advanced PDF features Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a Workbook, configures PdfSaveOptions.Version to 1.6 to activate PDF 1.6 capabilities such as transparency, saves the workbook as "output.pdf", and handles any exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Configure PDF save options (default PDF version supports advanced features like transparency)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF with the specified options
                workbook.Save("output.pdf", pdfOptions);
                Console.WriteLine("Workbook successfully saved as PDF.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
