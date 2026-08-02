// Title: Batch set a shared background image on the first worksheet of multiple Excel files and export to PDF with Aspose.Cells (.NET)
// Description: Loads a single image into memory, iterates over a list of .xlsx workbooks, assigns the image as the first sheet's background, and saves each workbook as a PDF while handling missing files and I/O errors.
// Keywords: Aspose.Cells | C# worksheet background image | batch Excel to PDF | shared image byte array | set worksheet background | export workbook to PDF | process multiple workbooks | Aspose.Cells .NET
// Common Searches: Aspose.Cells set worksheet background for multiple workbooks | batch convert Excel to PDF with background image C# | apply same background image to many Excel files using Aspose | load image once and reuse for Excel to PDF conversion | C# loop over workbooks set background and save as PDF
// Developer Intent: Programmatically apply one background image to the first sheet of each workbook in a collection and generate PDF files.
// Use Cases: Create branded PDF reports from a set of Excel templates by adding a company logo as a background on the first sheet before conversion. | Automate client‑specific workbooks where a watermark image is applied to the first sheet of every file and then exported to PDF. | Run a batch job that reads a single image file once, reuses it across many workbooks, and outputs PDFs to minimize I/O overhead.
// AI Prompts: Show how to modify the code to use a different background image for each workbook while still loading images only once. | Provide an example of processing a large number of workbooks asynchronously with progress reporting using Aspose.Cells. | Explain how to keep the original Excel files unchanged and save all generated PDFs to a separate output folder.

using System;
using System.IO;
using Aspose.Cells;

// Loads a single image into memory, iterates over a list of .xlsx workbooks, assigns the image as the first sheet's background, and saves each workbook as a PDF while handling missing files and I/O errors.
public class BatchBackgroundToPdf
{
    public static void Run()
    {
        // Path to the shared background image file
        string backgroundImagePath = "sharedBackground.jpg";

        // Verify background image exists
        if (!File.Exists(backgroundImagePath))
        {
            Console.WriteLine($"Background image not found: {backgroundImagePath}");
            return;
        }

        // Load the image once into a byte array
        byte[] backgroundData;
        try
        {
            backgroundData = File.ReadAllBytes(backgroundImagePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read background image: {ex.Message}");
            return;
        }

        // List of workbook files to process
        string[] workbookFiles = new string[]
        {
            "Workbook1.xlsx",
            "Workbook2.xlsx",
            "Workbook3.xlsx"
        };

        foreach (string wbPath in workbookFiles)
        {
            // Verify workbook file exists
            if (!File.Exists(wbPath))
            {
                Console.WriteLine($"Workbook file not found: {wbPath}");
                continue;
            }

            try
            {
                // Load the workbook from file
                Workbook workbook = new Workbook(wbPath);

                // Set the background image of the first worksheet
                Worksheet firstSheet = workbook.Worksheets[0];
                firstSheet.BackgroundImage = backgroundData;

                // Determine output PDF file name
                string pdfPath = Path.ChangeExtension(wbPath, ".pdf");

                // Save the workbook as PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Processed '{wbPath}' -> '{pdfPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{wbPath}': {ex.Message}");
            }
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            BatchBackgroundToPdf.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
