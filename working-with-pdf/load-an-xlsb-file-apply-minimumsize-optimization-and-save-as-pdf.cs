// Title: Convert an XLSB workbook to a minimized PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert XLSB to PDF with file size reduction | How to enable MinimizeSize option when saving PDF with Aspose.Cells | Using PdfSaveOptions OptimizationOptions to shrink PDF output in .NET | Load binary Excel (.xlsb) and export to PDF with size optimization in C# | Reflection based approach to set PDF optimization in Aspose.Cells
// Tags: Aspose.Cells PDF size optimization | XLSB to PDF conversion C# | PdfSaveOptions MinimizeSize property | OptimizationOptions in Aspose.Cells | Reflection set Aspose.Cells PDF options

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies the XLSB input file, loads it with Aspose.Cells, creates a PdfSaveOptions object, uses reflection to enable the MinimizeSize flag inside OptimizationOptions (if supported), and saves the workbook as a reduced‑size PDF.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Attempt to set MinimumSize optimization if the API supports it
            PropertyInfo optProp = typeof(PdfSaveOptions).GetProperty("OptimizationOptions");
            if (optProp != null && optProp.CanWrite)
            {
                // Create an instance of the optimization options class
                object optInstance = Activator.CreateInstance(optProp.PropertyType);
                PropertyInfo minSizeProp = optProp.PropertyType.GetProperty("MinimizeSize");
                if (minSizeProp != null && minSizeProp.CanWrite)
                {
                    // Enable size minimization
                    minSizeProp.SetValue(optInstance, true);
                }
                // Assign the configured optimization options back to PdfSaveOptions
                optProp.SetValue(pdfOptions, optInstance);
            }

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
