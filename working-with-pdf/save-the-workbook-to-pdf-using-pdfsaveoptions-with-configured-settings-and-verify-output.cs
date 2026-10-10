// Title: Export a C# Aspose.Cells workbook to a single‑page PDF with PdfSaveOptions and verify the generated file
// AI Prompts: Write C# code that creates a workbook, fills it with sample data, configures PdfSaveOptions with OnePagePerSheet set to true, and saves the workbook as a PDF file. | Add logic that checks whether the destination folder exists, creates it if necessary, then saves the PDF and confirms the file exists and its size is greater than zero. | Wrap the PDF conversion and directory handling in a try‑catch block that logs any exceptions that occur.
// Common Searches: c# aspnet export excel to pdf one page per sheet using aspose.cells | how to ensure output directory exists before saving pdf with aspose.cells in .net | verify pdf file size after saving workbook with aspose.cells pdfsaveoptions | asp.net core save workbook as pdf and check if file is not empty | aspocells pdfsaveoptions onepagepersheet example c#
// Tags: Aspose.Cells PDF export single page per sheet | C# workbook to PDF conversion | verify PDF file existence and size .NET | create output folder before saving PDF | PdfSaveOptions configuration for PDF output

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new workbook, populates it with sample data, configures PdfSaveOptions to fit each worksheet on a single PDF page, ensures the target directory exists, saves the workbook as SampleData.pdf, and then checks that the PDF file was created and is not empty, reporting its size or any errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["A3"].PutValue("Oranges");
            sheet.Cells["B3"].PutValue(200);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // Fit each sheet on a single PDF page
                // Compliance property omitted for compatibility with older Aspose.Cells versions
            };

            // Define the output PDF file path
            string pdfPath = "SampleData.pdf";

            // Ensure the directory for the PDF exists (handle cases where pdfPath has no directory part)
            string pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
            if (string.IsNullOrEmpty(pdfDir))
            {
                pdfDir = Directory.GetCurrentDirectory();
            }

            if (!Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook as PDF using the configured options
            workbook.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created and has content
            if (File.Exists(pdfPath))
            {
                FileInfo info = new FileInfo(pdfPath);
                if (info.Length > 0)
                {
                    Console.WriteLine($"PDF saved successfully. Size: {info.Length} bytes.");
                }
                else
                {
                    Console.WriteLine("PDF file was created but is empty.");
                }
            }
            else
            {
                Console.WriteLine("Failed to create the PDF file.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
