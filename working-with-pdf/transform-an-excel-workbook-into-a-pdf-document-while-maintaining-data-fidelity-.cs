using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class ExcelToPdfConverter
    {
        public static void ConvertExcelToPdf(string excelPath, string pdfPath)
        {
            // Load the Excel workbook (lifecycle rule: create/load)
            Workbook workbook = new Workbook(excelPath);

            // Ensure formulas are calculated to preserve data fidelity
            workbook.CalculateFormula();

            // Convert the workbook to PDF using the provided ConversionUtility rule
            // This method handles loading and saving internally, preserving layout.
            ConversionUtility.Convert(excelPath, pdfPath);

            Console.WriteLine($"Conversion completed: '{excelPath}' -> '{pdfPath}'");
        }

        // Example usage
        public static void Main()
        {
            string sourceExcel = "input.xlsx";   // Path to the source Excel file
            string destinationPdf = "output.pdf"; // Desired PDF output path

            ConvertExcelToPdf(sourceExcel, destinationPdf);
        }
    }
}