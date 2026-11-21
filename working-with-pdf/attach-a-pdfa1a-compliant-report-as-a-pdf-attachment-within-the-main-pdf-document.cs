// Title: Embed a PDF file as an OLE object in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: how to embed a PDF as an OLE object in an Excel sheet using Aspose.Cells C# | c# example for inserting PDF bytes into Excel workbook with Aspose.Cells | save an Excel file with a PDF attachment via Aspose.Cells .NET | add PDF file to worksheet as OLE shape using Aspose.Cells API | read PDF file and create OLE object in Excel with Aspose.Cells C#
// Tags: aspocells addoleobject pdf | c# embed pdf into excel worksheet | excel ole object insertion pdf | aspocells read pdf bytes shape | c# save workbook with embedded pdf | aspocells pdf attachment excel

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook, reads a PDF file into a byte array, embeds the PDF as an OLE object shape on the first worksheet using Aspose.Cells' AddOleObject method, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for the source Excel file, PDF to embed, and output workbook
            string mainExcelPath = "main.xlsx";
            string reportPdfPath = "report.pdf";
            string outputExcelPath = "main_with_attachment.xlsx";

            // Verify that required files exist
            if (!File.Exists(mainExcelPath))
            {
                Console.WriteLine($"Error: Excel file not found – {mainExcelPath}");
                return;
            }
            if (!File.Exists(reportPdfPath))
            {
                Console.WriteLine($"Error: PDF file not found – {reportPdfPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(mainExcelPath);

            // Use the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Read PDF bytes to embed as OLE object
            byte[] pdfData;
            try
            {
                pdfData = File.ReadAllBytes(reportPdfPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read PDF file: {ex.Message}");
                return;
            }

            // Embed the PDF as an OLE object in the worksheet
            // Parameters: upper-left row, column, height (pixels), width (pixels),
            // lower-right row, column, OLE object data (byte[])
            int upperLeftRow = 0;
            int upperLeftColumn = 0;
            int height = 200;
            int width = 300;
            int lowerRightRow = upperLeftRow;
            int lowerRightColumn = upperLeftColumn;

            sheet.Shapes.AddOleObject(upperLeftRow, upperLeftColumn, height, width,
                                      lowerRightRow, lowerRightColumn, pdfData);

            // Save the modified workbook
            workbook.Save(outputExcelPath);
            Console.WriteLine($"Workbook saved successfully: {outputExcelPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception: {ex.Message}");
        }
    }
}
