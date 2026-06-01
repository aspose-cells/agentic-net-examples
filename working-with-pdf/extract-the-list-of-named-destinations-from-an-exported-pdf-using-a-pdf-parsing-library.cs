using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Path to the Excel file that will be exported to PDF
            string workbookPath = "PdfBookmarkDestinationDemo.xlsx";

            // Verify the Excel file exists before loading
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(workbookPath);

            // List all worksheet names (demonstrates successful loading)
            Console.WriteLine("Worksheets in the workbook:");
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"- {sheet.Name}");
            }

            // Export the workbook to PDF (creates the PDF with bookmarks if any)
            string pdfPath = "PdfBookmarkDestinationDemo.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook exported to PDF: {pdfPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime errors and display a concise message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}