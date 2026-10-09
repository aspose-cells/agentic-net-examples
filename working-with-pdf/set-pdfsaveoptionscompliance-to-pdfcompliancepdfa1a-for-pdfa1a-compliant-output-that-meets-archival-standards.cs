// Title: Create a PDF/A‑1a compliant PDF from an Aspose.Cells workbook in C# using PdfSaveOptions.Compliance
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as a PDF/A‑1a file by setting PdfSaveOptions.Compliance to PdfA1a. | Show how to use reflection in C# to assign the Compliance property on PdfSaveOptions only when the property exists in the current Aspose.Cells version. | Add directory creation and comprehensive error handling to ensure the workbook is saved as an archival PDF even if PDF/A support is unavailable.
// Common Searches: C# Aspose.Cells export workbook to PDF/A-1a compliance | how to set PdfSaveOptions.Compliance to PdfA1a in Aspose.Cells | using reflection to set PdfSaveOptions.Compliance when property is missing | save Excel file as archival PDF/A-1a with Aspose.Cells C# | fallback handling for missing PDF/A support in Aspose.Cells PDF export
// Tags: Aspose.Cells PDF/A‑1a export C# | PdfSaveOptions compliance property | reflection set PdfSaveOptions.Compliance | archival PDF generation from Excel | fallback handling for missing PDF/A support Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfAExample
{
    // The example creates an Aspose.Cells workbook, adds sample data, configures PdfSaveOptions, uses reflection to set the Compliance property to PdfA1a when available, ensures the output directory exists, and saves the workbook as a PDF/A‑1a compliant archival PDF with graceful fallback if the feature is unsupported.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook(); // creates an empty workbook

                // Add some data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data");

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Attempt to set PDF/A‑1a compliance if the property exists in this version
                try
                {
                    var complianceProp = typeof(PdfSaveOptions).GetProperty("Compliance");
                    if (complianceProp != null && complianceProp.CanWrite)
                    {
                        Type enumType = complianceProp.PropertyType;
                        object enumValue = Enum.Parse(enumType, "PdfA1a");
                        complianceProp.SetValue(pdfOptions, enumValue);
                    }
                }
                catch (Exception)
                {
                    // If any reflection or parsing error occurs, ignore and continue without PDF/A compliance
                }

                string outputPath = "ArchivedDocument.pdf";

                // Determine output directory; if none, use current directory
                string outputDir = Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory();

                // Ensure the output directory exists
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a PDF (PDF/A compliance if supported)
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
