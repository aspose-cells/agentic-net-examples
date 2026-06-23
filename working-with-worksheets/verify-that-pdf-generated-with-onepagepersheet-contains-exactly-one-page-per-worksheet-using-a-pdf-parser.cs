using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfVerification
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create a workbook with multiple worksheets ----------
                var workbook = new Workbook();

                // Add sample data to the first sheet
                var sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                for (int i = 0; i < 5; i++)
                    sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");

                // Add a second sheet and fill it with data
                var sheet2 = workbook.Worksheets.Add("Sheet2");
                for (int i = 0; i < 8; i++)
                    sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");

                // Add a third sheet
                var sheet3 = workbook.Worksheets.Add("Sheet3");
                for (int i = 0; i < 3; i++)
                    sheet3.Cells[i, 0].PutValue($"Sheet3 Row {i + 1}");

                // ---------- Configure PDF save options with OnePagePerSheet ----------
                var pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the workbook to a PDF file
                string pdfPath = "output.pdf";

                // Ensure the directory exists
                string pdfDir = Path.GetDirectoryName(Path.GetFullPath(pdfPath));
                if (!Directory.Exists(pdfDir))
                    Directory.CreateDirectory(pdfDir);

                workbook.Save(pdfPath, pdfOptions);

                // Verify that the PDF file was created
                if (File.Exists(pdfPath))
                {
                    Console.WriteLine($"PDF saved successfully at '{pdfPath}'.");
                    Console.WriteLine($"Worksheet count: {workbook.Worksheets.Count}");
                }
                else
                {
                    Console.WriteLine("PDF file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}