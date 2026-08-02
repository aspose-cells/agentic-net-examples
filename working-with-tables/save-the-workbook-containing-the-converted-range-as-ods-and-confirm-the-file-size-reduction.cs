// Title: C# Aspose.Cells: Save Workbook as ODS and Verify Size Reduction
// Description: Creates a 1,000‑row by 10‑column workbook, saves it as XLSX, then uses OdsSaveOptions (LibreOffice generator) to export the same workbook to ODS, and finally compares the two file sizes to confirm whether the ODS file is smaller.
// Keywords: Aspose.Cells ODS conversion | C# save workbook as ODS | OdsSaveOptions LibreOffice | XLSX to ODS file size | reduce spreadsheet size | .NET Aspose.Cells export ODS | compare XLSX and ODS sizes | Aspose.Cells file size optimization
// Common Searches: How to export an Aspose.Cells workbook to ODS in C# | Aspose.Cells OdsSaveOptions example for size reduction | Convert large Excel file to ODS and check file size | C# compare XLSX and ODS file sizes using Aspose.Cells | LibreOffice generator impact on ODS output size
// Developer Intent: Save a workbook as ODS and determine if the ODS file is smaller than the original XLSX.
// Use Cases: Produce a compact ODS version of a massive Excel report for easier sharing. | Automate a routine that selects the most storage‑efficient format (XLSX vs ODS). | Leverage the LibreOffice generator to ensure maximum compatibility while minimizing file size.
// AI Prompts: Generate C# code with Aspose.Cells that converts a workbook to ODS using OdsSaveOptions and prints the size difference. | Create a reusable method that accepts an XLSX path, saves it as ODS, and returns the percentage reduction in bytes. | Explain how OdsGeneratorType.LibreOffice affects ODS compatibility and file size compared to the default generator.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// Creates a 1,000‑row by 10‑column workbook, saves it as XLSX, then uses OdsSaveOptions (LibreOffice generator) to export the same workbook to ODS, and finally compares the two file sizes to confirm whether the ODS file is smaller.
class Program
{
    static void Main()
    {
        // Create a new workbook and populate it with sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        for (int row = 0; row < 1000; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                worksheet.Cells[row, col].PutValue($"R{row}C{col}");
            }
        }

        // Save the workbook in its original XLSX format
        string xlsxPath = "original.xlsx";
        workbook.Save(xlsxPath, SaveFormat.Xlsx);

        // Save the same workbook as ODS using OdsSaveOptions
        string odsPath = "converted.ods";
        OdsSaveOptions odsOptions = new OdsSaveOptions();
        odsOptions.GeneratorType = OdsGeneratorType.LibreOffice; // optional setting
        workbook.Save(odsPath, odsOptions);

        // Compare file sizes to confirm reduction
        long xlsxSize = new FileInfo(xlsxPath).Length;
        long odsSize = new FileInfo(odsPath).Length;

        Console.WriteLine($"XLSX file size: {xlsxSize} bytes");
        Console.WriteLine($"ODS file size: {odsSize} bytes");

        if (odsSize < xlsxSize)
        {
            Console.WriteLine($"Size reduced by {xlsxSize - odsSize} bytes after conversion to ODS.");
        }
        else
        {
            Console.WriteLine("ODS file is not smaller than the original XLSX.");
        }
    }
}
