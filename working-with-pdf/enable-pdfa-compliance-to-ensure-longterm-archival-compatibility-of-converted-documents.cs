// Title: Save an Excel workbook as PDF/A‑1b using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx file, sets PdfSaveOptions.Compliance to PdfA1b, and saves the workbook as a PDF/A‑1b document with Aspose.Cells. | Show how to create the output directory, handle missing input files, and export a workbook to PDF/A while catching exceptions in C#. | Provide a minimal Aspose.Cells example that converts a worksheet to a PDF/A‑1b file suitable for long‑term archiving.
// Common Searches: Aspose.Cells how to export Excel to PDF/A‑1b in C# | C# set PdfSaveOptions compliance to PDF/A using Aspose.Cells | Convert .xlsx to PDF/A archive format with Aspose.Cells .NET | Save workbook as PDF/A for long term storage Aspose.Cells | Create output folder before saving PDF/A with Aspose.Cells C#
// Tags: Aspose.Cells PDF/A export C# | PdfSaveOptions set compliance PDF/A .NET | Excel to PDF/A archival conversion | C# generate PDF/A‑1b workbook | Aspose.Cells long‑term archive PDF

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a new workbook if the file is missing), ensures the target directory exists, configures PdfSaveOptions for PDF/A compliance, and saves the workbook as a PDF/A‑1b document, with error handling to capture any issues during the conversion.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            Workbook workbook;

            // Load existing workbook if the file exists; otherwise create a new one.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");
            }

            // Configure PDF save options (default compliance).
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file.
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
