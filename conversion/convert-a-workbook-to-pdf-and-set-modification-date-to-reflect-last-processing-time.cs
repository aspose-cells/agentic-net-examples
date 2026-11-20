// Title: Convert an Excel workbook to PDF and set the PDF modification date with Aspose.Cells for .NET
// Common Searches: how to export an .xlsx file to PDF and update its timestamp using Aspose.Cells in C# | Aspose.Cells C# example for converting Excel to PDF and setting the last modified date | C# code to save an Excel workbook as PDF and change the PDF file's last write time | using Aspose.Cells to convert a workbook to PDF and modify file metadata programmatically | set PDF file modification time after converting from Excel with Aspose.Cells .NET
// Tags: Aspose.Cells Excel to PDF conversion C# | set PDF file last write time programmatically | Workbook.Save using SaveFormat.Pdf Aspose | update file metadata after Excel to PDF export | create placeholder workbook if input file missing Aspose

using Aspose.Cells;
using System;
using System.IO;

// The sample loads (or creates) an Excel workbook, saves it as a PDF using Aspose.Cells' SaveFormat.Pdf, then updates the PDF file's last write time to the current processing time, with proper exception handling for Aspose.Cells-specific and general errors.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = "input.xlsx";

        // Path where the PDF will be saved
        string pdfPath = "output.pdf";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not
            if (!File.Exists(excelPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(excelPath);
            }

            // Load the workbook (no password supplied)
            Workbook workbook = new Workbook(excelPath);

            // Convert and save the workbook as a PDF document
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Update the PDF file's modification date to the current processing time
            DateTime processingTime = DateTime.Now;
            File.SetLastWriteTime(pdfPath, processingTime);
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protected files)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles all other runtime errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
