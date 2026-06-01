using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Define file paths
            string sourcePath = "sample.xlsx";
            string pdfPath = "sample.pdf";

            // Create a new workbook and add sample data
            Workbook workbook = new Workbook(); // uses Workbook() constructor rule
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);

            // Save the workbook as XLSX using the Save(string, SaveFormat) rule
            workbook.Save(sourcePath, SaveFormat.Xlsx);

            // Convert the XLSX file to PDF using default settings via ConversionUtility
            ConversionUtility.Convert(sourcePath, pdfPath); // uses Convert(string, string) rule

            // Verify that the PDF file was created and output its size
            if (File.Exists(pdfPath))
            {
                FileInfo pdfInfo = new FileInfo(pdfPath);
                Console.WriteLine($"PDF conversion successful. File size: {pdfInfo.Length} bytes.");
            }
            else
            {
                Console.WriteLine("PDF conversion failed. Output file not found.");
            }

            // Clean up temporary XLSX file (optional)
            if (File.Exists(sourcePath))
            {
                File.Delete(sourcePath);
            }
        }
    }
}