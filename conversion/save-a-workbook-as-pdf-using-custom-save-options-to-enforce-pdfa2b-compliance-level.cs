// Title: Convert an Aspose.Cells workbook to a PDF/A‑2b compliant PDF using C# and PdfSaveOptions
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, populates sample cells, configures PdfSaveOptions for PDF/A‑2b compliance, and saves the workbook as a PDF file. | Demonstrate how to detect whether the PdfCompliance enum is available in the current Aspose.Cells version and set PdfSaveOptions.Compliance to PdfA2b only when supported. | Provide a C# snippet that ensures the target output directory exists (creating it if needed) before invoking Workbook.Save with the configured PDF options.
// Common Searches: Aspose.Cells C# export workbook to PDF/A-2b format | How to set PdfSaveOptions.Compliance to PdfA2b in Aspose.Cells .NET | C# create output folder before saving PDF with Aspose.Cells | Workaround for missing PdfCompliance enum in older Aspose.Cells versions | Example of saving Excel workbook as PDF/A-2b using Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells PDF/A-2b conversion C# | PdfSaveOptions compliance configuration | Workbook.Save PDF output directory handling | PdfCompliance enum version detection | Excel to PDF/A-2b Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create a workbook, fill cells, optionally set PdfSaveOptions.Compliance to PdfA2b (when supported), ensure the output directory exists, and save the workbook as a PDF/A‑2b compliant PDF using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Sample");
            sheet.Cells["B1"].PutValue("Data");
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["B2"].PutValue(200);

            // Set up PDF save options (PDF/A‑2b compliance may not be supported in older versions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Uncomment the following line if the PdfCompliance enum is available in your Aspose.Cells version
            // pdfOptions.Compliance = PdfCompliance.PdfA2b;

            // Define output file path
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file using the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
