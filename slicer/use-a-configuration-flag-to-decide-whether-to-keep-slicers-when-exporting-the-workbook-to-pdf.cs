// Title: Use a boolean flag to toggle slicer visibility when saving an Excel workbook to PDF with Aspose.Cells for .NET
// Common Searches: how to export Excel to PDF with slicers hidden using Aspose.Cells C# | Aspose.Cells PDF export show or hide slicers based on a configuration flag | set SlicerRenderingMode property in PdfSaveOptions via reflection in C# | conditionally render slicers when converting a workbook to PDF with Aspose.Cells | C# code to keep slicers visible in PDF output using Aspose.Cells
// Tags: PdfSaveOptions SlicerRenderingMode property | conditional slicer visibility Aspose.Cells | export workbook to PDF with slicer control | reflection set enum value C# Aspose.Cells | hide slicers in PDF output using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, creates PdfSaveOptions, and uses reflection to set the SlicerRenderingMode enum to Show or Hide based on a boolean flag, then saves the workbook as a PDF while optionally preserving slicer visibility.
    class Program
    {
        static void Main(string[] args)
        {
            // Configuration flag: set to true to keep slicers in the PDF, false to hide them
            bool keepSlicers = true;

            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Create PDF save options
                var pdfOptions = new PdfSaveOptions();

                // NOTE: In some older versions of Aspose.Cells the SlicerRenderingMode property
                // is not available. If it exists, set it according to the configuration flag.
                // This block uses reflection to avoid compile‑time errors on versions lacking the property.
                var slicerProp = typeof(PdfSaveOptions).GetProperty("SlicerRenderingMode");
                if (slicerProp != null && slicerProp.CanWrite)
                {
                    // Determine the enum type (Aspose.Cells.Rendering.SlicerRenderingMode)
                    var enumType = slicerProp.PropertyType;
                    var enumValue = Enum.Parse(enumType, keepSlicers ? "Show" : "Hide");
                    slicerProp.SetValue(pdfOptions, enumValue);
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the workbook to PDF with the specified options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
