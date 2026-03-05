using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsXlsxDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle: create)
            Workbook newWorkbook = new Workbook();

            // Add some sample data to the default worksheet
            Worksheet sheet = newWorkbook.Worksheets[0];
            sheet.Name = "SampleData";
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue(999.99);
            sheet.Cells["A3"].PutValue("Phone");
            sheet.Cells["B3"].PutValue(699.99);

            // 2. Save the newly created workbook as XLSX (lifecycle: save)
            string newFilePath = "CreatedSample.xlsx";
            newWorkbook.Save(newFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Created workbook saved to '{newFilePath}'.");

            // 3. Load an existing XLSX file (lifecycle: load)
            string existingFilePath = "ExistingData.xlsx"; // replace with actual path
            Workbook loadedWorkbook = new Workbook(existingFilePath);

            // Perform a simple operation: add a new row at the end
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            int lastRow = loadedSheet.Cells.MaxDataRow + 1;
            loadedSheet.Cells[lastRow, 0].PutValue("Tablet");
            loadedSheet.Cells[lastRow, 1].PutValue(399.99);

            // 4. Save the modified workbook back to XLSX
            string modifiedFilePath = "ModifiedData.xlsx";
            loadedWorkbook.Save(modifiedFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Modified workbook saved to '{modifiedFilePath}'.");

            // 5. Convert the XLSX file to PDF using ConversionUtility (optional demonstration)
            string pdfOutputPath = "ConvertedToPdf.pdf";
            ConversionUtility.Convert(modifiedFilePath, pdfOutputPath);
            Console.WriteLine($"Workbook converted to PDF at '{pdfOutputPath}'.");
        }
    }
}