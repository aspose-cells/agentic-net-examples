// Title: Add a shared PNG background to the first worksheet of multiple Excel files and export each workbook to PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set same background image on first sheet of several workbooks | How to batch convert Excel workbooks to PDF while applying a common worksheet picture using Aspose.Cells | Aspose.Cells C# add picture to worksheet before saving as PDF in a loop | Processing multiple .xlsx files to add background PNG and export to PDF with Aspose.Cells
// Tags: add background picture to worksheet Aspose.Cells | batch process workbooks C# Aspose.Cells | export workbook to PDF Aspose.Cells | Pictures.Add method usage Aspose.Cells | load multiple Excel files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Iterates over a list of Excel workbooks, inserts a shared PNG image onto the first worksheet of each using the Pictures.Add API, and saves the modified workbooks as PDF files via Aspose.Cells.
class BatchWorkbookProcessor
{
    static void Main()
    {
        // Paths of the workbooks to process
        string[] workbookPaths = new string[]
        {
            @"C:\Data\Workbook1.xlsx",
            @"C:\Data\Workbook2.xlsx",
            // add more paths as needed
        };

        // Path of the background image
        string backgroundImagePath = @"C:\Images\Background.png";

        // Verify that the background image exists before processing
        if (!File.Exists(backgroundImagePath))
        {
            Console.WriteLine($"Background image not found: {backgroundImagePath}");
            return;
        }

        foreach (string wbPath in workbookPaths)
        {
            try
            {
                // Verify workbook file exists
                if (!File.Exists(wbPath))
                {
                    Console.WriteLine($"Workbook not found: {wbPath}");
                    continue;
                }

                // Load the workbook from file
                Workbook workbook = new Workbook(wbPath);

                // Access the first worksheet (index 0)
                Worksheet firstSheet = workbook.Worksheets[0];

                // Add the background image as a picture covering the sheet.
                // This avoids the need for System.Drawing.Image.
                firstSheet.Pictures.Add(0, 0, backgroundImagePath);

                // Define the output PDF path (same name, .pdf extension)
                string pdfPath = Path.ChangeExtension(wbPath, ".pdf");

                // Export the workbook to PDF
                workbook.Save(pdfPath, SaveFormat.Pdf);

                Console.WriteLine($"Processed and saved PDF: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook '{wbPath}': {ex.Message}");
            }
        }
    }
}
