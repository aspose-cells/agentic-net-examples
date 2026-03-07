using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class RenderSlicerToPdf
{
    static void Main()
    {
        // Path to the source XLSX file that contains a slicer
        string sourcePath = "input.xlsx";

        // Desired output PDF file path
        string destPath = "output.pdf";

        // Convert the Excel workbook (including slicers) to PDF.
        // The ConversionUtility.Convert method handles loading and saving internally.
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Workbook with slicer has been successfully converted to PDF.");
    }
}