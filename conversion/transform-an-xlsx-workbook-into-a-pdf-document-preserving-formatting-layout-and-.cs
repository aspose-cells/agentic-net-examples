using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Utility;      // For ConversionUtility (optional)

class XlsxToPdfConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired output PDF file path
        string destPath = "output.pdf";

        // Load the workbook from the XLSX file (load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Save the workbook as PDF, preserving formatting, layout and pagination (save rule)
        workbook.Save(destPath, SaveFormat.Pdf);

        // Alternative using ConversionUtility (also valid):
        // ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}