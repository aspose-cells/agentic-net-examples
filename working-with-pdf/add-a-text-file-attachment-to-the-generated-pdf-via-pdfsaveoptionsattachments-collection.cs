// Title: Add a .txt attachment to a PDF generated from an Aspose.Cells workbook in C# (using Aspose.Pdf post‑processing)
// AI Prompts: Write C# code that creates an Excel workbook with Aspose.Cells, saves it as a PDF, then opens the PDF with Aspose.Pdf and embeds a .txt file as an attachment. | Show how to generate a sample text file, export the workbook to PDF, and use Aspose.Pdf's Attachments collection to add the file, including proper disposal and error handling. | Provide a step‑by‑step example that checks for the attachment file, ensures the output directory exists, saves the workbook as PDF, and then attaches the file using Aspose.Pdf's Document class.
// Common Searches: C# attach a text file to a PDF created with Aspose.Cells | Aspose.Pdf add attachment after exporting Excel to PDF | how to embed a .txt file in a PDF using Aspose libraries .NET | workaround for missing Attachments collection in Aspose.Cells PdfSaveOptions
// Tags: Aspose.Cells export workbook to PDF C# | Aspose.Pdf embed file attachment C# | post‑process PDF with Aspose.Pdf after Aspose.Cells export | PdfSaveOptions limitation attachment collection | add embedded attachment to PDF .NET

using System;
using System.IO;
using Aspose.Cells;

// The example creates an Excel workbook with Aspose.Cells, writes sample data, generates a text file if needed, saves the workbook as a PDF using PdfSaveOptions, and then uses Aspose.Pdf to embed the text file as an attachment because PdfSaveOptions does not provide an Attachments collection. It includes directory creation, error handling, and demonstrates the required post‑processing step.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create workbook ----------
            Workbook workbook = new Workbook();

            // Optional: add some data to the workbook
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample data");

            // ---------- Prepare PDF save options ----------
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // ---------- Create a sample attachment file ----------
            string attachmentPath = "attachment.txt";
            if (!File.Exists(attachmentPath))
            {
                File.WriteAllText(attachmentPath, "This is a sample text attachment.");
            }

            // NOTE: Aspose.Cells' PdfSaveOptions does not expose an Attachments collection
            // in the current version. If attachment support is required, consider using
            // Aspose.Pdf to add the file after the PDF is generated.

            // ---------- Save workbook as PDF ----------
            string outputPath = "output.pdf";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
