using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Author: Aspose.Cells .NET example – convert XLSX to PDF and check file size
class Program
{
    static void Main()
    {
        // Define source Excel file and destination PDF file paths
        string sourcePath = "input.xlsx";
        string pdfPath = "output.pdf";

        // Convert the Excel workbook to PDF using default settings
        // (ConversionUtility.Convert is the provided save rule)
        ConversionUtility.Convert(sourcePath, pdfPath);

        // Verify that the PDF was created and output its file size
        FileInfo pdfInfo = new FileInfo(pdfPath);
        if (pdfInfo.Exists)
        {
            Console.WriteLine($"PDF saved successfully. Size: {pdfInfo.Length} bytes");
        }
        else
        {
            Console.WriteLine("PDF conversion failed: output file not found.");
        }
    }
}