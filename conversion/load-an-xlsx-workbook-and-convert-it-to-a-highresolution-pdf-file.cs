using System;
using Aspose.Cells;               // Core Aspose.Cells namespace
using Aspose.Cells.Utility;      // For ConversionUtility if needed (not used here)

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Path where the high‑resolution PDF will be saved
            string pdfPath = "output.pdf";

            // Load the workbook from the XLSX file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Save the loaded workbook as a PDF file (uses Workbook.Save(string, SaveFormat) method)
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook '{sourcePath}' successfully converted to PDF at '{pdfPath}'.");
        }
    }
}