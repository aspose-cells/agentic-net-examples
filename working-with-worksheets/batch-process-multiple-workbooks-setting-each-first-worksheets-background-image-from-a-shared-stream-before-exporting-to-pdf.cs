using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering; // PdfSaveOptions and SheetSet

class BatchBackgroundToPdf
{
    static void Main()
    {
        // Folder containing source Excel files
        string inputFolder = @"C:\InputWorkbooks";
        // Folder where PDF files will be saved
        string outputFolder = @"C:\OutputPdfs";
        // Path to the shared background image
        string backgroundImagePath = @"C:\SharedResources\background.jpg";

        try
        {
            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify background image exists
            if (!File.Exists(backgroundImagePath))
            {
                Console.WriteLine($"Background image not found: {backgroundImagePath}");
                return;
            }

            // Load the background image once into a byte array
            byte[] backgroundBytes;
            using (FileStream imgStream = new FileStream(backgroundImagePath, FileMode.Open, FileAccess.Read))
            using (MemoryStream ms = new MemoryStream())
            {
                imgStream.CopyTo(ms);
                backgroundBytes = ms.ToArray();
            }

            // Process each Excel workbook in the input folder
            foreach (string excelPath in Directory.GetFiles(inputFolder, "*.xlsx"))
            {
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Workbook not found (skipped): {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook from file
                    Workbook workbook = new Workbook(excelPath);

                    // Set the background image of the first worksheet
                    Worksheet firstSheet = workbook.Worksheets[0];
                    firstSheet.BackgroundImage = backgroundBytes;

                    // Prepare PDF save options
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        SheetSet = SheetSet.Visible // render all visible sheets
                    };

                    // Determine the output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(excelPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the workbook as PDF
                    workbook.Save(pdfPath, pdfOptions);

                    Console.WriteLine($"Processed '{excelPath}' -> '{pdfPath}'");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}